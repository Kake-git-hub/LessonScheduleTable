import type { Assignment, SessionData } from '../types'
import { personKey } from './schedule'
import { constraintFor, hasAvailability, isStudentAvailable, isRegularLessonPair } from './constraints'
import { canTeachSubject, teachableBaseSubjects, BASE_SUBJECTS } from './subjects'
import {
  getSlotNumber,
  getStudentSubject,
  countTeacherLoad,
  getTeacherAssignedDates,
  getTeacherSlotNumbersOnDate,
  getTeacherPrevSlotStudentIds,
  countStudentSlotsOnDate,
  getStudentSlotNumbersOnDate,
  countStudentAssignedDates,
  countStudentLoad,
  countStudentSubjectLoad,
  assignmentSignature,
  hasMeaningfulManualAssignment,
  getTeacherStudentSlotsOnDate,
  getStudentSubjectsOnAdjacentSlots,
} from './assignments'

export interface ChangeLogEntry {
  slot: string
  action: string
  detail: string
}

/** Yield helper — resolves on next macrotask so the UI stays responsive. */
const yieldToMain = (): Promise<void> => new Promise((r) => setTimeout(r, 0))

export const buildIncrementalAutoAssignments = async (
  data: SessionData,
  slots: string[],
): Promise<{ assignments: Record<string, Assignment[]>; changeLog: ChangeLogEntry[]; changedPairSignatures: Record<string, string[]>; addedPairSignatures: Record<string, string[]>; changeDetails: Record<string, Record<string, string>> }> => {
  const changeLog: ChangeLogEntry[] = []
  const changedPairSigSetBySlot: Record<string, Set<string>> = {}
  const addedPairSigSetBySlot: Record<string, Set<string>> = {}
  const changeDetailsBySlot: Record<string, Record<string, string>> = {}

  // --- Helpers for detailed submission-based reasons ---
  const lastAutoAt = data.settings.lastAutoAssignedAt ?? 0

  /** Describe what changed in a student's submission since last auto-assign */
  const describeStudentSubmissionChange = (studentId: string): string => {
    const student = data.students.find((s) => s.id === studentId)
    if (!student) return ''
    const recentEntries = (data.submissionLog ?? [])
      .filter((e) => e.personId === studentId && e.personType === 'student' && e.submittedAt > lastAutoAt)
      .sort((a, b) => b.submittedAt - a.submittedAt)
    if (recentEntries.length === 0) return ''

    const latest = recentEntries[0]
    const timeStr = new Date(latest.submittedAt).toLocaleDateString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    if (latest.type === 'initial') {
      const slotsDetail = latest.subjectSlots
        ? Object.entries(latest.subjectSlots).map(([s, c]) => `${s}${c}コマ`).join(', ')
        : ''
      return `${student.name}: 前回コマ割り後に希望を新規提出(${timeStr})${slotsDetail ? ` [${slotsDetail}]` : ''}`
    }
    const prevEntries = (data.submissionLog ?? [])
      .filter((e) => e.personId === studentId && e.personType === 'student' && e.submittedAt <= lastAutoAt)
      .sort((a, b) => b.submittedAt - a.submittedAt)
    const prev = prevEntries.length > 0 ? prevEntries[0] : null
    const changes: string[] = []
    if (latest.subjectSlots && prev?.subjectSlots) {
      const allSubjs = new Set([...Object.keys(latest.subjectSlots), ...Object.keys(prev.subjectSlots)])
      for (const subj of allSubjs) {
        const oldVal = prev.subjectSlots[subj] ?? 0
        const newVal = latest.subjectSlots[subj] ?? 0
        if (oldVal !== newVal) changes.push(`${subj}: ${oldVal}→${newVal}コマ`)
      }
    }
    if (latest.unavailableSlots && prev?.unavailableSlots) {
      const oldCount = prev.unavailableSlots.length
      const newCount = latest.unavailableSlots.length
      if (oldCount !== newCount) changes.push(`不可コマ数: ${oldCount}→${newCount}`)
    }
    const diffStr = changes.length > 0 ? ` [${changes.join(', ')}]` : ''
    return `${student.name}: 前回コマ割り後に希望を変更(${timeStr})${diffStr}`
  }

  /** Describe what changed in a teacher's submission since last auto-assign */
  const describeTeacherSubmissionChange = (teacherId: string): string => {
    const teacher = data.teachers.find((t) => t.id === teacherId)
    if (!teacher) return ''
    const recentEntries = (data.submissionLog ?? [])
      .filter((e) => e.personId === teacherId && e.personType === 'teacher' && e.submittedAt > lastAutoAt)
      .sort((a, b) => b.submittedAt - a.submittedAt)
    if (recentEntries.length === 0) return ''
    const latest = recentEntries[0]
    const timeStr = new Date(latest.submittedAt).toLocaleDateString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    if (latest.type === 'initial') {
      return `${teacher.name}: 前回コマ割り後に出勤希望を新規提出(${timeStr})`
    }
    return `${teacher.name}: 前回コマ割り後に出勤希望を変更(${timeStr})`
  }

  const markChangedPair = (slot: string, assignment: Assignment, detail: string): void => {
    if (assignment.isRegular) return
    if (!hasMeaningfulManualAssignment(assignment)) return
    if (!changedPairSigSetBySlot[slot]) changedPairSigSetBySlot[slot] = new Set<string>()
    const sig = assignmentSignature(assignment)
    changedPairSigSetBySlot[slot].add(sig)
    if (!changeDetailsBySlot[slot]) changeDetailsBySlot[slot] = {}
    const prev = changeDetailsBySlot[slot][sig]
    changeDetailsBySlot[slot][sig] = prev ? `${prev}\n${detail}` : detail
  }
  const markAddedPair = (slot: string, assignment: Assignment, detail?: string): void => {
    if (assignment.isRegular) return
    if (!hasMeaningfulManualAssignment(assignment)) return
    if (!addedPairSigSetBySlot[slot]) addedPairSigSetBySlot[slot] = new Set<string>()
    const sig = assignmentSignature(assignment)
    addedPairSigSetBySlot[slot].add(sig)
    if (detail) {
      if (!changeDetailsBySlot[slot]) changeDetailsBySlot[slot] = {}
      changeDetailsBySlot[slot][sig] = detail
    }
  }
  const teacherIds = new Set(data.teachers.map((t) => t.id))
  const studentIds = new Set(data.students.map((s) => s.id))
  const result: Record<string, Assignment[]> = {}

  // Pre-populate result with actual results (recorded slots) so student load counting includes them
  if (data.actualResults) {
    for (const [slot, results] of Object.entries(data.actualResults)) {
      result[slot] = results.map((r) => ({
        teacherId: r.teacherId,
        studentIds: [...r.studentIds],
        subject: r.subject,
        studentSubjects: r.studentSubjects ? { ...r.studentSubjects } : undefined,
      }))
    }
  }

  // Build submission order map: earlier initial submission → higher priority (lower rank number)
  const submissionOrderMap = new Map<string, number>()
  if (data.submissionLog) {
    let rank = 0
    for (const entry of data.submissionLog) {
      if (entry.type === 'initial' && entry.personType === 'student' && !submissionOrderMap.has(entry.personId)) {
        submissionOrderMap.set(entry.personId, rank++)
      }
    }
  }
  // Students who haven't submitted get lowest priority
  const maxRank = submissionOrderMap.size

  // Phase 1: Clean up existing assignments — handle deleted teachers/students (skip regular lessons)
  for (const slot of slots) {
    const existing = data.assignments[slot]
    if (!existing || existing.length === 0) continue

    // Preserve regular lesson assignments as-is
    if (existing.every((a) => a.isRegular)) {
      result[slot] = [...existing]
      continue
    }

    const cleaned: Assignment[] = []
    for (const assignment of existing) {
      // Keep regular assignments untouched
      if (assignment.isRegular) {
        cleaned.push(assignment)
        continue
      }

      // Check teacher still exists
      if (!teacherIds.has(assignment.teacherId)) {
        const usedTeachers = new Set(cleaned.map((a) => a.teacherId))
        // Build student-subject pairs for grade-aware replacement check
        const studentSubjectPairs = assignment.studentIds.map(sid => {
          const s = data.students.find(st => st.id === sid)
          const subj = assignment.studentSubjects?.[sid] ?? assignment.subject
          return { grade: s?.grade ?? '', subject: subj }
        }).filter(p => p.grade && p.subject)
        const replacement = data.teachers.find((t) => {
          if (usedTeachers.has(t.id)) return false
          if (!hasAvailability(data.availability, 'teacher', t.id, slot)) return false
          return studentSubjectPairs.every(({ grade, subject }) => canTeachSubject(t.subjects, grade, subject))
        })
        if (replacement) {
          const changedAssignment = { ...assignment, teacherId: replacement.id }
          cleaned.push(changedAssignment)
          markChangedPair(slot, changedAssignment, `講師差替: ${replacement.name} に変更（元の講師が削除済）`)
          changeLog.push({ slot, action: '講師差替', detail: `${replacement.name} に変更（元の講師が削除済）` })
        } else {
          changeLog.push({ slot, action: '講師削除', detail: `割当解除（講師が削除済・代替不可）` })
        }
        continue
      }

      // Check teacher still has availability for this slot
      if (!hasAvailability(data.availability, 'teacher', assignment.teacherId, slot)) {
        const teacherName = data.teachers.find((t) => t.id === assignment.teacherId)?.name ?? '?'
        const teacherChangeInfo = describeTeacherSubmissionChange(assignment.teacherId)
        const usedTeachers = new Set(cleaned.map((a) => a.teacherId))
        const studentSubjectPairs2 = assignment.studentIds.map(sid => {
          const s = data.students.find(st => st.id === sid)
          const subj = assignment.studentSubjects?.[sid] ?? assignment.subject
          return { grade: s?.grade ?? '', subject: subj }
        }).filter(p => p.grade && p.subject)
        const replacement = data.teachers.find((t) => {
          if (usedTeachers.has(t.id)) return false
          if (!hasAvailability(data.availability, 'teacher', t.id, slot)) return false
          return studentSubjectPairs2.every(({ grade, subject }) => canTeachSubject(t.subjects, grade, subject))
        })
        if (replacement) {
          const changedAssignment = { ...assignment, teacherId: replacement.id }
          cleaned.push(changedAssignment)
          const reason = teacherChangeInfo || `${teacherName}がこのコマの希望を取り消したため`
          markChangedPair(slot, changedAssignment, `講師差替: ${teacherName} → ${replacement.name}\n理由: ${reason}`)
          changeLog.push({ slot, action: '講師差替', detail: `${teacherName} → ${replacement.name}（希望取消のため）` })
        } else {
          changeLog.push({ slot, action: '割当解除', detail: `${teacherName} の希望が取り消されたため解除` })
        }
        continue
      }

      // Check students still exist AND are available for this slot
      const validStudentIds = assignment.studentIds.filter((sid) => {
        if (!studentIds.has(sid)) return false
        const student = data.students.find((s) => s.id === sid)
        if (!student) return false
        return isStudentAvailable(student, slot)
      })
      const removedStudentIds = assignment.studentIds.filter((sid) => !validStudentIds.includes(sid))
      for (const sid of removedStudentIds) {
        const student = data.students.find((s) => s.id === sid)
        const studentName = student?.name ?? `ID:${sid}`
        if (!studentIds.has(sid)) {
          changeLog.push({ slot, action: '生徒削除', detail: `${studentName} を解除（削除済）` })
        } else {
          const changeInfo = describeStudentSubmissionChange(sid)
          const reason = changeInfo || `${studentName}がこのコマを不可に変更したため`
          changeLog.push({ slot, action: '生徒解除', detail: `${studentName} を解除\n理由: ${reason}` })
        }
      }

      if (validStudentIds.length > 0) {
        const changedAssignment = { ...assignment, studentIds: validStudentIds }
        cleaned.push(changedAssignment)
        if (assignment.studentIds.length !== validStudentIds.length) {
          const removedNames = removedStudentIds.map((sid) => {
            const name = data.students.find((s) => s.id === sid)?.name ?? sid
            const changeInfo = describeStudentSubmissionChange(sid)
            return changeInfo || `${name}: このコマを不可に変更`
          }).join('\n')
          markChangedPair(slot, changedAssignment, `生徒解除:\n${removedNames}`)
        }
      } else if (removedStudentIds.length > 0) {
        const changedAssignment = { ...assignment, studentIds: [] }
        cleaned.push(changedAssignment)
        changeLog.push({ slot, action: '生徒全員解除', detail: `講師のみ残留（予定変更のため）` })
      } else {
        cleaned.push(assignment)
      }
    }

    if (cleaned.length > 0) {
      result[slot] = cleaned
    }
  }

  // Phase 1.5: Remove excess assignments when requested slots were reduced
  const specialLoadMap = new Map<string, number>()
  for (const slot of slots) {
    const slotAssignments = result[slot] ?? []
    for (const assignment of slotAssignments) {
      if (assignment.isRegular) continue
      for (const studentId of assignment.studentIds) {
        const subj = getStudentSubject(assignment, studentId)
        const key = `${studentId}|${subj}`
        specialLoadMap.set(key, (specialLoadMap.get(key) ?? 0) + 1)
      }
    }
  }

  const reverseSlots = [...slots].reverse()
  for (const slot of reverseSlots) {
    const slotAssignments = result[slot]
    if (!slotAssignments || slotAssignments.length === 0) continue

    for (const assignment of slotAssignments) {
      if (assignment.isRegular || assignment.studentIds.length === 0) continue

      const remainingStudentIds: string[] = []
      let removedAny = false
      for (const studentId of assignment.studentIds) {
        const student = data.students.find((s) => s.id === studentId)
        const subj = getStudentSubject(assignment, studentId)
        const requested = student?.subjectSlots[subj] ?? 0
        const key = `${studentId}|${subj}`
        const currentLoad = specialLoadMap.get(key) ?? 0

        if (currentLoad > requested) {
          specialLoadMap.set(key, currentLoad - 1)
          removedAny = true
          const changeInfo = describeStudentSubmissionChange(studentId)
          const reason = changeInfo || `${student?.name ?? studentId}が${subj}の希望コマ数を${requested}コマに減らしたため`
          changeLog.push({ slot, action: '希望減で解除', detail: `${student?.name ?? studentId} (${subj}) を解除\n理由: ${reason}` })
          continue
        }
        remainingStudentIds.push(studentId)
      }

      if (removedAny) {
        assignment.studentIds = remainingStudentIds
        if (remainingStudentIds.length > 0) {
          markChangedPair(slot, assignment, `希望コマ数減少により一部生徒を解除`)
        }
      }
    }
  }

  // Phase 2: Fill empty student positions in existing non-regular assignments
  for (const slot of slots) {
    if (!result[slot] || result[slot].length === 0) continue
    const slotAssignments = result[slot]
    for (let idx = 0; idx < slotAssignments.length; idx++) {
      const assignment = slotAssignments[idx]
      if (assignment.isRegular) continue
      if (assignment.studentIds.length >= 2) continue
      if (!assignment.teacherId) continue

      const teacher = data.teachers.find((t) => t.id === assignment.teacherId)
      if (!teacher) continue

      const usedStudentIdsInSlot = new Set(slotAssignments.flatMap((a) => a.studentIds))

      const candidates = data.students.filter((student) => {
        if (usedStudentIdsInSlot.has(student.id)) return false
        if (!isStudentAvailable(student, slot)) return false
        if (constraintFor(data.constraints, teacher.id, student.id) === 'incompatible') return false
        // Check student-student constraints with existing students in this assignment
        if (assignment.studentIds.some((existingSid) => constraintFor(data.constraints, existingSid, student.id) === 'incompatible')) return false
        // Student must be able to learn at least one subject the teacher can teach (grade-aware), with remaining demand
        return student.subjects.some((baseSubj) => {
          if (!canTeachSubject(teacher.subjects, student.grade, baseSubj)) return false
          const requested = student.subjectSlots[baseSubj] ?? 0
          const allocated = countStudentSubjectLoad(result, student.id, baseSubj)
          return allocated < requested
        })
      })

      if (candidates.length > 0 && assignment.studentIds.length < 2) {
        const best = candidates.sort((a, b) => {
          const aRem = Object.values(a.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, a.id)
          const bRem = Object.values(b.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, b.id)
          return bRem - aRem
        })[0]
        // Pick the best viable subject for this new student
        const bestSubj = teachableBaseSubjects(teacher.subjects, best.grade).find((baseSubj) => {
          if (!best.subjects.includes(baseSubj)) return false
          const requested = best.subjectSlots[baseSubj] ?? 0
          const allocated = countStudentSubjectLoad(result, best.id, baseSubj)
          return allocated < requested
        }) ?? assignment.subject
        // Reconstruct studentSubjects
        const studentSubjects: Record<string, string> = {}
        for (const sid of assignment.studentIds) {
          studentSubjects[sid] = getStudentSubject(assignment, sid)
        }
        studentSubjects[best.id] = bestSubj
        assignment.studentIds = [...assignment.studentIds, best.id]
        assignment.studentSubjects = studentSubjects
        const addChangeInfo = describeStudentSubmissionChange(best.id)
        const addReason = addChangeInfo || `${best.name}の${bestSubj}が未充足(残${(best.subjectSlots[bestSubj] ?? 0) - countStudentSubjectLoad(result, best.id, bestSubj)}コマ)のため空き枠に追加`
        markChangedPair(slot, assignment, `生徒追加: ${best.name}(${bestSubj})\n理由: ${addReason}`)
        changeLog.push({ slot, action: '生徒追加', detail: `${best.name}(${bestSubj}) を追加` })
      }
    }
  }

  // Phase 3: Fill empty slots — minimize total teacher attendance dates, distribute students evenly

  // Compute date ordering for first-half bias
  const allDatesInOrder: string[] = []
  for (const s of slots) {
    const d = s.split('_')[0]
    if (allDatesInOrder.length === 0 || allDatesInOrder[allDatesInOrder.length - 1] !== d) {
      allDatesInOrder.push(d)
    }
  }
  const totalDates = allDatesInOrder.length
  const dateIndexMap = new Map<string, number>()
  for (let i = 0; i < totalDates; i++) dateIndexMap.set(allDatesInOrder[i], i)

  const deskCountLimit = data.settings.deskCount ?? 0

  let slotCounter = 0
  for (const slot of slots) {
    // Yield every 5 slots to keep the page responsive
    if (++slotCounter % 5 === 0) await yieldToMain()

    const currentDate = slot.split('_')[0]
    const currentSlotNum = getSlotNumber(slot)

    // Initialize from existing assignments (allow adding more pairs to a slot)
    const existingAssignments = result[slot] ?? []
    const slotAssignments: Assignment[] = [...existingAssignments]
    const usedTeacherIdsInSlot = new Set<string>(existingAssignments.map((a) => a.teacherId))
    const usedStudentIdsInSlot = new Set<string>(existingAssignments.flatMap((a) => a.studentIds))
    // Regular-only slots are allowed — new pairs can be added alongside regular lessons
    // Skip slots at desk count limit
    if (deskCountLimit > 0 && slotAssignments.length >= deskCountLimit) {
      continue
    }

    // First-half bias: earlier dates get a bonus (max +25 for first date, 0 for last)
    const dateIdx = dateIndexMap.get(currentDate) ?? 0
    const firstHalfBonus = totalDates > 1 ? Math.round(25 * (1 - dateIdx / (totalDates - 1))) : 0

    const teachers = data.teachers.filter((teacher) =>
      hasAvailability(data.availability, 'teacher', teacher.id, slot),
    )

    // Sort teachers: strongly prefer those already assigned on this date (minimize total attendance days)
    const sortedTeachers = [...teachers].sort((a, b) => {
      const aDates = getTeacherAssignedDates(result, a.id)
      const bDates = getTeacherAssignedDates(result, b.id)
      const aOnDate = aDates.has(currentDate) ? 0 : 1
      const bOnDate = bDates.has(currentDate) ? 0 : 1
      if (aOnDate !== bOnDate) return aOnDate - bOnDate
      // Then prefer teachers with fewer total attendance dates
      return aDates.size - bDates.size
    })

    for (const teacher of sortedTeachers) {
      if (usedTeacherIdsInSlot.has(teacher.id)) continue
      // Stop if desk count limit reached
      if (deskCountLimit > 0 && slotAssignments.length >= deskCountLimit) break

      const candidates = data.students.filter((student) => {
        if (usedStudentIdsInSlot.has(student.id)) return false
        if (!isStudentAvailable(student, slot)) return false
        if (constraintFor(data.constraints, teacher.id, student.id) === 'incompatible') return false
        return student.subjects.some((baseSubj) => {
          return canTeachSubject(teacher.subjects, student.grade, baseSubj)
        })
      })

      if (candidates.length === 0) continue

      const teacherDates = getTeacherAssignedDates(result, teacher.id)
      const isExistingDate = teacherDates.has(currentDate)
      const teacherLoad = countTeacherLoad(result, teacher.id)

      // Teacher consecutive slot bonus
      const teacherSlotsOnDate = getTeacherSlotNumbersOnDate(result, teacher.id, currentDate)
      const teacherIsConsecutive = teacherSlotsOnDate.some((n) => Math.abs(n - currentSlotNum) === 1)
      const teacherConsecutiveBonus = teacherIsConsecutive ? 20 : 0

      // Students that the teacher taught in the immediately previous slot (avoid same student consecutive)
      const prevSlotStudentIds = new Set(getTeacherPrevSlotStudentIds(result, teacher.id, currentDate, currentSlotNum))

      let bestPlan: { score: number; assignment: Assignment } | null = null

      for (const combo of [...candidates.map((s) => [s]), ...candidates.flatMap((l, i) => candidates.slice(i + 1).map((r) => [l, r]))]) {
        // Avoid assigning the same student to this teacher's consecutive slot
        const hasSameStudentConsecutive = combo.some((st) => prevSlotStudentIds.has(st.id))
        if (hasSameStudentConsecutive) continue

        // Check student-student constraints within this combo
        if (combo.length === 2 && constraintFor(data.constraints, combo[0].id, combo[1].id) === 'incompatible') continue

        // --- Determine subject assignment (same or mixed) ---
        // Find base subjects ALL students in combo can learn and teacher can teach to ALL of them
        const commonBaseSubjects = (BASE_SUBJECTS as readonly string[]).filter((baseSubj) =>
          combo.every((student) => student.subjects.includes(baseSubj) && canTeachSubject(teacher.subjects, student.grade, baseSubj)),
        )
        const viableCommonSubjects = commonBaseSubjects.filter((baseSubj) =>
          combo.every((student) => {
            const requested = student.subjectSlots[baseSubj] ?? 0
            const allocated = countStudentSubjectLoad(result, student.id, baseSubj)
            return allocated < requested
          }),
        )

        // For 2-student combos: also try mixed-subject pairing
        type SubjectPlan = { isMixed: false; subject: string } | { isMixed: true; studentSubjects: Record<string, string>; primarySubject: string }
        const subjectPlans: SubjectPlan[] = []

        // Add same-subject plans
        for (const subj of viableCommonSubjects) {
          subjectPlans.push({ isMixed: false, subject: subj })
        }

        // Add mixed-subject plans for 2-student combos
        if (combo.length === 2) {
          const [s1, s2] = combo
          const s1Viable = teachableBaseSubjects(teacher.subjects, s1.grade).filter((baseSubj) => {
            if (!s1.subjects.includes(baseSubj)) return false
            const req = s1.subjectSlots[baseSubj] ?? 0
            const alloc = countStudentSubjectLoad(result, s1.id, baseSubj)
            return alloc < req
          })
          const s2Viable = teachableBaseSubjects(teacher.subjects, s2.grade).filter((baseSubj) => {
            if (!s2.subjects.includes(baseSubj)) return false
            const req = s2.subjectSlots[baseSubj] ?? 0
            const alloc = countStudentSubjectLoad(result, s2.id, baseSubj)
            return alloc < req
          })
          // Only add mixed plans where subjects actually differ
          for (const subj1 of s1Viable) {
            for (const subj2 of s2Viable) {
              if (subj1 === subj2) continue
              subjectPlans.push({
                isMixed: true,
                studentSubjects: { [s1.id]: subj1, [s2.id]: subj2 },
                primarySubject: subj1,
              })
            }
          }
        }

        if (subjectPlans.length === 0) continue

        for (const plan of subjectPlans) {
        // --- Student distribution scoring ---
        let studentScore = 0
        for (const st of combo) {
          const slotsOnDate = countStudentSlotsOnDate(result, st.id, currentDate)
          const existingSlotNums = getStudentSlotNumbersOnDate(result, st.id, currentDate)

          // Penalty for same-day multiple slots (avoid if possible)
          if (slotsOnDate > 0) {
            studentScore -= 60
            const isConsecutive = existingSlotNums.some(
              (n) => Math.abs(n - currentSlotNum) === 1,
            )
            if (isConsecutive) {
              studentScore += 50
            } else {
              studentScore -= 30
            }
          }

          // Prefer students with more remaining slots (even distribution)
          const totalRequested = Object.values(st.subjectSlots).reduce((s, c) => s + c, 0)
          const totalAssigned = countStudentLoad(result, st.id)
          studentScore += (totalRequested - totalAssigned) * 10

          // Prefer students with fewer assigned dates (spread across days)
          const assignedDates = countStudentAssignedDates(result, st.id)
          studentScore -= assignedDates * 5

          // Submission order bonus: earlier submitters get priority (max +15)
          const submissionRank = submissionOrderMap.get(st.id) ?? maxRank
          studentScore += Math.max(0, 15 - submissionRank * 2)
        }

        // Regular lesson pair bonus: prefer assigning regular-lesson teacher-student combos
        const regularPairBonus = combo.reduce((s, st) =>
          s + (isRegularLessonPair(data.regularLessons, teacher.id, st.id) ? 30 : 0), 0)

        // Same-day same-pair consecutive bonus
        let pairConsecutiveBonus = 0
        for (const st of combo) {
          const existingPairSlots = getTeacherStudentSlotsOnDate(result, teacher.id, st.id, currentDate)
          if (existingPairSlots.length > 0) {
            const isConsecutive = existingPairSlots.some((n) => Math.abs(n - currentSlotNum) === 1)
            pairConsecutiveBonus += isConsecutive ? 60 : -40
          }
        }

        // Mixed-subject penalty: same subject pairs are slightly preferred
        const mixedSubjectPenalty = plan.isMixed ? -15 : 0

        // Consecutive same-subject penalty: avoid same student having the same subject in adjacent slots
        let consecutiveSameSubjectPenalty = 0
        for (const st of combo) {
          const subj = plan.isMixed ? (plan.studentSubjects[st.id] ?? '') : plan.subject
          const adjacentSubjects = getStudentSubjectsOnAdjacentSlots(result, st.id, currentDate, currentSlotNum)
          if (adjacentSubjects.includes(subj)) {
            consecutiveSameSubjectPenalty -= 20
          }
        }

        const score = 100 +
          (isExistingDate ? 80 : -50) +
          teacherConsecutiveBonus +
          firstHalfBonus +
          regularPairBonus +
          pairConsecutiveBonus +
          (combo.length === 2 ? 30 : 0) +
          mixedSubjectPenalty +
          consecutiveSameSubjectPenalty +
          studentScore -
          teacherLoad * 2

        if (!bestPlan || score > bestPlan.score) {
          const assignment: Assignment = plan.isMixed
            ? { teacherId: teacher.id, studentIds: combo.map((s) => s.id), subject: plan.primarySubject, studentSubjects: plan.studentSubjects }
            : { teacherId: teacher.id, studentIds: combo.map((s) => s.id), subject: plan.subject }
          bestPlan = { score, assignment }
        }
        } // end for plan
      }

      if (bestPlan) {
        slotAssignments.push(bestPlan.assignment)
        usedTeacherIdsInSlot.add(teacher.id)
        for (const sid of bestPlan.assignment.studentIds) usedStudentIdsInSlot.add(sid)
      }
    }

    if (slotAssignments.length > existingAssignments.length) {
      result[slot] = slotAssignments
      // Only log newly added assignments
      for (const a of slotAssignments.slice(existingAssignments.length)) {
        const tName = data.teachers.find((t) => t.id === a.teacherId)?.name ?? '?'
        const tChange = describeTeacherSubmissionChange(a.teacherId)
        const sNames = a.studentIds.map((sid) => {
          const name = data.students.find((s) => s.id === sid)?.name ?? '?'
          const subj = getStudentSubject(a, sid)
          return `${name}(${subj})`
        }).join(', ')
        const sChanges = a.studentIds
          .map((sid) => describeStudentSubmissionChange(sid))
          .filter(Boolean)
          .join(' / ')
        const detailParts = [`新規割当: ${tName} × ${sNames}`]
        if (tChange) detailParts.push(`[講師] ${tChange}`)
        if (sChanges) detailParts.push(`[生徒] ${sChanges}`)
        const fullDetail = detailParts.join(' | ')
        markAddedPair(slot, a, fullDetail)
        changeLog.push({ slot, action: '新規割当', detail: fullDetail })
      }
    }
  }

  const changedPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(changedPairSigSetBySlot)) {
    const addedSet = addedPairSigSetBySlot[slot] ?? new Set<string>()
    const filtered = [...signatureSet].filter((sig) => !addedSet.has(sig))
    if (filtered.length > 0) changedPairSignatures[slot] = filtered
  }

  const addedPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(addedPairSigSetBySlot)) {
    if (signatureSet.size > 0) addedPairSignatures[slot] = [...signatureSet]
  }

  return { assignments: result, changeLog, changedPairSignatures, addedPairSignatures, changeDetails: changeDetailsBySlot }
}

/** Mendan (interview) FCFS auto-assign: each parent gets exactly 1 slot with 1 manager */
export const buildMendanAutoAssignments = (
  data: SessionData,
  slots: string[],
): { assignments: Record<string, Assignment[]>; unassignedParents: string[] } => {
  // Get managers and their availability
  const managerAvailability = new Map<string, Set<string>>()
  for (const manager of (data.managers ?? [])) {
    const key = personKey('manager', manager.id)
    managerAvailability.set(manager.id, new Set(data.availability[key] ?? []))
  }

  // Get parents sorted by submittedAt (FCFS) — only submitted parents
  const sortedParents = data.students
    .filter((s) => s.submittedAt > 0)
    .sort((a, b) => a.submittedAt - b.submittedAt)

  const result: Record<string, Assignment[]> = {}
  // Copy existing non-regular assignments only
  for (const slot of slots) {
    const existing = data.assignments[slot]
    if (existing?.length) {
      const nonRegular = existing.filter((a) => !a.isRegular)
      if (nonRegular.length > 0) result[slot] = [...nonRegular]
    }
  }

  // Track which parents are already assigned (ignore regular lesson assignments)
  const assignedParents = new Set<string>()
  for (const slot of slots) {
    for (const a of (result[slot] ?? [])) {
      if (a.isRegular) continue
      for (const sid of a.studentIds) assignedParents.add(sid)
    }
  }
  // Also count parents in recorded actual results as already assigned
  if (data.actualResults) {
    for (const results of Object.values(data.actualResults)) {
      for (const r of results) {
        for (const sid of r.studentIds) assignedParents.add(sid)
      }
    }
  }

  const unassignedParents: string[] = []

  for (const parent of sortedParents) {
    if (assignedParents.has(parent.id)) continue

    const parentKey = personKey('student', parent.id)
    const parentSlots = new Set(data.availability[parentKey] ?? [])
    if (parentSlots.size === 0) {
      unassignedParents.push(parent.name)
      continue
    }

    let assigned = false
    for (const slot of slots) {
      if (!parentSlots.has(slot)) continue

      const slotAssignments = result[slot] ?? []
      const usedManagers = new Set(slotAssignments.map((a) => a.teacherId))
      const usedStudents = new Set(slotAssignments.flatMap((a) => a.studentIds))

      if (usedStudents.has(parent.id)) continue

      // Check desk count
      const deskCount = data.settings.deskCount ?? 0
      if (deskCount > 0 && slotAssignments.length >= deskCount) continue

      // Find available manager for this slot
      for (const [managerId, mSlots] of managerAvailability) {
        if (!mSlots.has(slot)) continue
        if (usedManagers.has(managerId)) continue

        // Assign!
        const assignment: Assignment = {
          teacherId: managerId,
          studentIds: [parent.id],
          subject: '面談',
        }
        result[slot] = [...(result[slot] ?? []), assignment]
        assignedParents.add(parent.id)
        assigned = true
        break
      }

      if (assigned) break
    }

    if (!assigned) {
      unassignedParents.push(parent.name)
    }
  }

  return { assignments: result, unassignedParents }
}
