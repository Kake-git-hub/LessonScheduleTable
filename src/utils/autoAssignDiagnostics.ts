/**
 * Auto-assign diagnostics: analyse why students remain unassigned after Phase 3.
 *
 * Intentionally avoids importing from autoAssign.ts to prevent circular dependencies.
 * The teacher-availability check is reproduced locally (same logic as
 * hasTeacherAvailabilityForAutoAssign in autoAssign.ts).
 */

import type { Assignment, SessionData, Student } from '../types'
import { constraintFor, hasAvailability, isStudentAvailable } from './constraints'
import { canTeachSubject } from './subjects'
import { evaluateConstraintCards } from './slotConstraints'
import { countStudentSubjectLoad, getIsoDayOfWeek, getSlotNumber } from './assignments'
import { slotLabel } from './schedule'

// ── Types ──────────────────────────────────────────────────────────────────

export type StudentDiagnostic = {
  studentId: string
  studentName: string
  grade: string
  /** Per-subject demand/fulfilment breakdown */
  subjectDemand: Record<string, { requested: number; assigned: number; remaining: number }>
  /** Total slots in the session (slot 0 excluded) */
  totalSlots: number
  /** Slots where the student is available */
  candidateSlots: number
  /** Counts of why the student could not be placed in each slot */
  blockReasons: {
    studentUnavailable: number
    noTeacherAvailable: number
    allTeachersIncompatible: number
    subjectMismatch: number
    constraintCardBlocked: number
    deskLimitReached: number
    alreadyAssignedInSlot: number
    demandAlreadyMet: number
  }
  /** The block reason with the highest count */
  primaryBottleneck: string
  /** Per-slot detail (for JSON export) */
  slotDetails: SlotDiagnosticDetail[]
}

export type SlotDiagnosticDetail = {
  slot: string
  slotLabel: string
  blocked: boolean
  reason: string
  detail?: string
}

// ── Internal helper ────────────────────────────────────────────────────────

/**
 * Replicates the same check as `hasTeacherAvailabilityForAutoAssign` in autoAssign.ts.
 * Kept local to avoid a circular import.
 */
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

// ── Main diagnostic function ───────────────────────────────────────────────

/**
 * After `buildIncrementalAutoAssignments` completes, call this function to
 * obtain per-student explanations for any remaining unmet demand.
 *
 * @param data    The session data (students, teachers, constraints, settings…)
 * @param result  The assignment map produced by the auto-assign run
 * @param slots   The ordered slot keys that were processed (same list passed to auto-assign)
 */
export const diagnoseUnassignedStudents = (
  data: SessionData,
  result: Record<string, Assignment[]>,
  slots: string[],
): StudentDiagnostic[] => {
  const deskCountLimit = data.settings.deskCount ?? 0
  const isMendan = data.settings.sessionType === 'mendan'
  const mendanStart = data.settings.mendanStartHour ?? 10

  // Filter slots to only those with slot number > 0 (exclude 午前 group slots)
  const analysisSlots = slots.filter((slot) => getSlotNumber(slot) > 0)

  const diagnostics: StudentDiagnostic[] = []

  for (const student of data.students) {
    // Only diagnose students with remaining unmet demand
    const subjectDemand: StudentDiagnostic['subjectDemand'] = {}
    let totalRemaining = 0

    for (const [subj, requested] of Object.entries(student.subjectSlots ?? {})) {
      const assigned = countStudentSubjectLoad(result, student.id, subj)
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
      demandAlreadyMet: 0,
    }

    const slotDetails: SlotDiagnosticDetail[] = []
    let candidateSlots = 0

    for (const slot of analysisSlots) {
      const slotAssignments = result[slot] ?? []
      const label = slotLabel(slot, isMendan, mendanStart)

      // 1. Student unavailable
      if (!isStudentAvailable(student, slot)) {
        blockReasons.studentUnavailable++
        slotDetails.push({ slot, slotLabel: label, blocked: true, reason: 'studentUnavailable' })
        continue
      }

      candidateSlots++

      // 2. Student already assigned in this slot
      const alreadyInSlot = slotAssignments.some((a) => a.studentIds.includes(student.id))
      if (alreadyInSlot) {
        blockReasons.alreadyAssignedInSlot++
        slotDetails.push({ slot, slotLabel: label, blocked: true, reason: 'alreadyAssignedInSlot' })
        continue
      }

      // 3. Desk count limit
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

      // 4–7. Teacher-based checks
      const availableTeachers = data.teachers.filter((t) => teacherAvailableForSlot(data, t.id, slot))

      if (availableTeachers.length === 0) {
        blockReasons.noTeacherAvailable++
        slotDetails.push({ slot, slotLabel: label, blocked: true, reason: 'noTeacherAvailable' })
        continue
      }

      // 5. All available teachers are incompatible
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

      // 6. Subject mismatch — none of the compatible teachers can teach any of the student's subjects
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

      // 7. Constraint card blocked for ALL subject-matching teachers
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

      // 8. All checks passed but demand is already met at this point in analysis
      // (Demand is counted from final result; if we reach here, it means a feasible
      //  placement existed but the student's demand was satisfied by other slots.)
      blockReasons.demandAlreadyMet++
      slotDetails.push({ slot, slotLabel: label, blocked: false, reason: 'demandAlreadyMet' })
    }

    // Determine primary bottleneck
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

// ── Label helpers ──────────────────────────────────────────────────────────

const BLOCK_REASON_LABELS: Record<string, string> = {
  studentUnavailable: '生徒が不可コマ/不可日',
  noTeacherAvailable: '出勤可能な講師がいない',
  allTeachersIncompatible: '全講師が相性不可',
  subjectMismatch: '科目が合う講師がいない',
  constraintCardBlocked: '制約カードでブロック',
  deskLimitReached: '机数上限',
  alreadyAssignedInSlot: 'そのスロットに既に割当済',
  demandAlreadyMet: '需要充足済（割当可能枠あり）',
}

export const blockReasonLabel = (reason: string): string =>
  BLOCK_REASON_LABELS[reason] ?? reason
