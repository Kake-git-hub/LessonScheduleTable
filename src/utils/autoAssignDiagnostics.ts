/**
 * 自動割当の残コマ診断: Phase 3 完了後に未割当の生徒がいる原因を分析する。
 *
 * autoAssign.ts からのインポートは循環依存を避けるため行わない。
 * 講師出勤判定は autoAssign.ts の hasTeacherAvailabilityForAutoAssign と同じロジックをローカルで再現。
 */

import type { Assignment, SessionData } from '../types'
import { constraintFor, hasAvailability, isStudentAvailable } from './constraints'
import { canTeachSubject } from './subjects'
import { evaluateConstraintCards } from './slotConstraints'
import { countStudentSubjectLoad, getIsoDayOfWeek, getSlotNumber } from './assignments'
import { slotLabel } from './schedule'

export type StudentDiagnostic = {
  studentId: string
  studentName: string
  grade: string
  subjectDemand: Record<string, { requested: number; assigned: number; remaining: number }>
  totalSlots: number
  candidateSlots: number
  blockReasons: {
    studentUnavailable: number
    noTeacherAvailable: number
    allTeachersIncompatible: number
    subjectMismatch: number
    constraintCardBlocked: number
    deskLimitReached: number
    alreadyAssignedInSlot: number
    allTeachersOccupied: number
    noRemainingSubjectTeacher: number
    demandAlreadyMet: number
  }
  primaryBottleneck: string
  slotDetails: SlotDiagnosticDetail[]
}

export type SlotDiagnosticDetail = {
  slot: string
  slotLabel: string
  blocked: boolean
  reason: string
  detail?: string
}

const teacherAvailableForSlot = (data: SessionData, teacherId: string, slot: string): boolean => {
  if (hasAvailability(data.availability, 'teacher', teacherId, slot)) return true
  if ((data.teacherSubmittedAt?.[teacherId] ?? 0) > 0) return false
  const [date] = slot.split('_')
  const dayOfWeek = getIsoDayOfWeek(date)
  const slotNumber = getSlotNumber(slot)
  return data.regularLessons.some(
    (lesson) =>
      lesson.teacherId === teacherId &&
      lesson.dayOfWeek === dayOfWeek &&
      lesson.slotNumber === slotNumber,
  )
}

export const diagnoseUnassignedStudents = (
  data: SessionData,
  result: Record<string, Assignment[]>,
  slots: string[],
): StudentDiagnostic[] => {
  const deskCountLimit = data.settings.deskCount ?? 0
  const isMendan = data.settings.sessionType === 'mendan'
  const mendanStart = data.settings.mendanStartHour ?? 10
  const analysisSlots = slots.filter((slot) => getSlotNumber(slot) > 0)
  const diagnostics: StudentDiagnostic[] = []

  for (const student of data.students) {
    const subjectDemand: StudentDiagnostic['subjectDemand'] = {}
    let totalRemaining = 0

    for (const [subj, requested] of Object.entries(student.subjectSlots ?? {})) {
      const assigned = countStudentSubjectLoad(result, student.id, subj, data.regularLessons)
      const remaining = Math.max(0, requested - assigned)
      subjectDemand[subj] = { requested, assigned, remaining }
      totalRemaining += remaining
    }

    if (totalRemaining === 0) continue

    const blockReasons: StudentDiagnostic['blockReasons'] = {
      studentUnavailable: 0,
      noTeacherAvailable: 0,
      allTeachersIncompatible: 0,
      subjectMismatch: 0,
      constraintCardBlocked: 0,
      deskLimitReached: 0,
      alreadyAssignedInSlot: 0,
      allTeachersOccupied: 0,
      noRemainingSubjectTeacher: 0,
      demandAlreadyMet: 0,
    }

    const slotDetails: SlotDiagnosticDetail[] = []
    let candidateSlots = 0

    for (const slot of analysisSlots) {
      const slotAssignments = result[slot] ?? []
      const label = slotLabel(slot, isMendan, mendanStart)

      if (!isStudentAvailable(student, slot)) {
        blockReasons.studentUnavailable++
        slotDetails.push({ slot, slotLabel: label, blocked: true, reason: 'studentUnavailable' })
        continue
      }

      candidateSlots++

      const alreadyInSlot = slotAssignments.some((a) => a.studentIds.includes(student.id))
      if (alreadyInSlot) {
        blockReasons.alreadyAssignedInSlot++
        slotDetails.push({ slot, slotLabel: label, blocked: true, reason: 'alreadyAssignedInSlot' })
        continue
      }

      if (deskCountLimit > 0 && slotAssignments.length >= deskCountLimit) {
        blockReasons.deskLimitReached++
        slotDetails.push({
          slot,
          slotLabel: label,
          blocked: true,
          reason: 'deskLimitReached',
          detail: `机数上限 ${deskCountLimit}`,
        })
        continue
      }

      const availableTeachers = data.teachers.filter((t) => teacherAvailableForSlot(data, t.id, slot))
      if (availableTeachers.length === 0) {
        blockReasons.noTeacherAvailable++
        slotDetails.push({ slot, slotLabel: label, blocked: true, reason: 'noTeacherAvailable' })
        continue
      }

      const compatibleTeachers = availableTeachers.filter(
        (t) => constraintFor(data.constraints, t.id, student.id) !== 'incompatible',
      )
      if (compatibleTeachers.length === 0) {
        blockReasons.allTeachersIncompatible++
        slotDetails.push({
          slot,
          slotLabel: label,
          blocked: true,
          reason: 'allTeachersIncompatible',
          detail: `出勤可能な全${availableTeachers.length}名の講師が相性不可`,
        })
        continue
      }

      const subjectMatchTeachers = compatibleTeachers.filter((t) =>
        student.subjects.some((subj) => canTeachSubject(t.subjects, student.grade, subj)),
      )
      if (subjectMatchTeachers.length === 0) {
        blockReasons.subjectMismatch++
        slotDetails.push({
          slot,
          slotLabel: label,
          blocked: true,
          reason: 'subjectMismatch',
          detail: `相性可の${compatibleTeachers.length}名全員が科目対応不可`,
        })
        continue
      }

      const groupLessonsOnDate = (() => {
        const [date] = slot.split('_')
        const dayOfWeek = getIsoDayOfWeek(date)
        return (data.groupLessons ?? []).filter((gl) => gl.dayOfWeek === dayOfWeek)
      })()

      const notCardBlocked = subjectMatchTeachers.filter((t) => {
        const evalResult = evaluateConstraintCards(
          student,
          slot,
          result,
          data.settings.slotsPerDay,
          data.regularLessons,
          groupLessonsOnDate,
          t.id,
        )
        return !evalResult.blocked
      })
      if (notCardBlocked.length === 0) {
        blockReasons.constraintCardBlocked++
        slotDetails.push({
          slot,
          slotLabel: label,
          blocked: true,
          reason: 'constraintCardBlocked',
          detail: `科目可の${subjectMatchTeachers.length}名全員が制約カードでブロック`,
        })
        continue
      }

      // 適合する全講師がこのコマで他の生徒に割当済かチェック
      const usedTeacherIdsInSlot = new Set(slotAssignments.map((a) => a.teacherId))
      const freeTeachers = notCardBlocked.filter((t) => !usedTeacherIdsInSlot.has(t.id))
      if (freeTeachers.length === 0) {
        blockReasons.allTeachersOccupied++
        slotDetails.push({
          slot,
          slotLabel: label,
          blocked: true,
          reason: 'allTeachersOccupied',
          detail: `適合する${notCardBlocked.length}名の講師が全員このコマで他の生徒に割当済`,
        })
        continue
      }

      // 空き講師が残需要科目を教えられるかチェック
      const remainingSubjects = Object.entries(subjectDemand)
        .filter(([, v]) => v.remaining > 0)
        .map(([subj]) => subj)
      const hasRemainingSubjectTeacher = freeTeachers.some((t) =>
        remainingSubjects.some((subj) => canTeachSubject(t.subjects, student.grade, subj)),
      )
      if (!hasRemainingSubjectTeacher) {
        blockReasons.noRemainingSubjectTeacher++
        slotDetails.push({
          slot,
          slotLabel: label,
          blocked: true,
          reason: 'noRemainingSubjectTeacher',
          detail: `空き講師${freeTeachers.length}名が残需要科目(${remainingSubjects.join(',')})に対応不可`,
        })
        continue
      }

      blockReasons.demandAlreadyMet++
      slotDetails.push({ slot, slotLabel: label, blocked: false, reason: 'demandAlreadyMet' })
    }

    const reasonEntries = Object.entries(blockReasons) as [keyof StudentDiagnostic['blockReasons'], number][]
    const maxEntry = reasonEntries.reduce<{ key: keyof StudentDiagnostic['blockReasons']; count: number } | null>(
      (best, [key, count]) => (count > 0 && (best === null || count > best.count) ? { key, count } : best),
      null,
    )
    const primaryBottleneck = maxEntry?.key ?? 'unknown'

    diagnostics.push({
      studentId: student.id,
      studentName: student.name,
      grade: student.grade,
      subjectDemand,
      totalSlots: analysisSlots.length,
      candidateSlots,
      blockReasons,
      primaryBottleneck,
      slotDetails,
    })
  }

  return diagnostics
}

const BLOCK_REASON_LABELS: Record<string, string> = {
  studentUnavailable: '生徒が不可コマ/不可日',
  noTeacherAvailable: '出勤可能な講師がいない',
  allTeachersIncompatible: '全講師が相性不可',
  subjectMismatch: '科目が合う講師がいない',
  constraintCardBlocked: '制約カードでブロック',
  deskLimitReached: '机数上限',
  alreadyAssignedInSlot: 'そのスロットに既に割当済',
  allTeachersOccupied: '適合講師が他の生徒に割当済',
  noRemainingSubjectTeacher: '残需要科目に対応できる空き講師なし',
  demandAlreadyMet: '割当可能だが他の生徒が優先された',
}

export const blockReasonLabel = (reason: string): string => BLOCK_REASON_LABELS[reason] ?? reason
