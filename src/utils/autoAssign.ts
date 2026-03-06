import type { Assignment, SessionData } from '../types'
import { personKey } from './schedule'
import { constraintFor, hasAvailability, isStudentAvailable } from './constraints'
import { canTeachSubject, teachableBaseSubjects, BASE_SUBJECTS } from './subjects'
import { evaluateConstraintCards, getDefaultConstraintCards } from './slotConstraints'
import {
  getSlotNumber,
  getIsoDayOfWeek,
  getStudentSubject,
  countTeacherLoad,
  getTeacherAssignedDates,
  getTeacherSlotNumbersOnDate,
  getTeacherPrevSlotStudentIds,
  countStudentSlotsOnDate,
  countStudentLoad,
  countStudentSubjectLoad,
  assignmentSignature,
  hasMeaningfulManualAssignment,
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
  onProgress?: (ratio: number) => void,
): Promise<{ assignments: Record<string, Assignment[]>; changeLog: ChangeLogEntry[]; changedPairSignatures: Record<string, string[]>; addedPairSignatures: Record<string, string[]>; makeupPairSignatures: Record<string, string[]>; changeDetails: Record<string, Record<string, string>> }> => {
  const changeLog: ChangeLogEntry[] = []
  const changedPairSigSetBySlot: Record<string, Set<string>> = {}
  const addedPairSigSetBySlot: Record<string, Set<string>> = {}
  const makeupPairSigSetBySlot: Record<string, Set<string>> = {}
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
  const markMakeupPair = (slot: string, assignment: Assignment, detail?: string): void => {
    if (assignment.isRegular) return
    if (!makeupPairSigSetBySlot[slot]) makeupPairSigSetBySlot[slot] = new Set<string>()
    const sig = assignmentSignature(assignment)
    makeupPairSigSetBySlot[slot].add(sig)
    if (detail) {
      if (!changeDetailsBySlot[slot]) changeDetailsBySlot[slot] = {}
      changeDetailsBySlot[slot][sig] = detail
    }
  }
  const teacherIds = new Set(data.teachers.map((t) => t.id))
  const studentIds = new Set(data.students.map((s) => s.id))
  const result: Record<string, Assignment[]> = {}

  // Pre-compute makeup student info: for each regular-lesson student, which regular slots are unavailable
  // Stores per-student array of { teacherId, dayOfWeek, slotNumber, date, subject, absentDate? } for each unavailable regular slot
  // Also includes students who were removed from actual results (absent in reality) — both regular and makeup
  // absentDate is set for actual-result-based absences so makeup is only placed on later dates
  type MakeupInfo = { teacherId: string; dayOfWeek: number; slotNumber: number; date: string; subject: string; absentDate?: string }
  const makeupStudentInfo = new Map<string, MakeupInfo[]>()
  // Build a set of ALL slot keys (including recorded) for regular-lesson matching
  const allSlotKeys = [...new Set([...slots, ...Object.keys(data.actualResults ?? {})])]
  for (const rl of data.regularLessons) {
    for (const sid of rl.studentIds) {
      const student = data.students.find((s) => s.id === sid)
      if (!student) continue
      const rlSubject = rl.studentSubjects?.[sid] ?? rl.subject
      for (const slot of allSlotKeys) {
        const [date] = slot.split('_')
        const dow = getIsoDayOfWeek(date)
        if (dow === rl.dayOfWeek && getSlotNumber(slot) === rl.slotNumber) {
          // Case 1: student is unavailable for this slot
          const needsMakeup = !isStudentAvailable(student, slot)
          // Case 2: slot has actual results recorded and student was removed (absent)
          const actualForSlot = data.actualResults?.[slot]
          const absentFromActual = actualForSlot != null
            && data.assignments[slot]?.some((a) => a.isRegular && a.studentIds.includes(sid))
            && !actualForSlot.some((r) => r.studentIds.includes(sid))
          if (needsMakeup || absentFromActual) {
            const arr = makeupStudentInfo.get(sid) ?? []
            arr.push({ teacherId: rl.teacherId, dayOfWeek: rl.dayOfWeek, slotNumber: rl.slotNumber, date, subject: rlSubject, ...(absentFromActual ? { absentDate: date } : {}) })
            makeupStudentInfo.set(sid, arr)
          }
        }
      }
    }
  }

  // Case 3: Detect makeup assignments (non-regular with regularMakeupInfo) where student was absent from actual results
  // These generate new makeup demand, just like regular lesson absences
  if (data.actualResults) {
    for (const [slot, actualResults] of Object.entries(data.actualResults)) {
      const origAssignments = data.assignments[slot] ?? []
      for (const orig of origAssignments) {
        if (orig.isRegular || !orig.regularMakeupInfo) continue
        for (const sid of orig.studentIds) {
          if (!orig.regularMakeupInfo[sid]) continue // only students who were makeup
          // Check if this student is absent from actual results in this slot
          const studentInActual = actualResults.some((r) => r.studentIds.includes(sid))
          if (studentInActual) continue
          // Student was absent from their makeup slot — generate new makeup demand
          const [absentDate] = slot.split('_')
          const mkInfo = orig.regularMakeupInfo[sid]
          const subject = orig.studentSubjects?.[sid] ?? orig.subject
          const arr = makeupStudentInfo.get(sid) ?? []
          arr.push({ teacherId: orig.teacherId, dayOfWeek: mkInfo.dayOfWeek, slotNumber: mkInfo.slotNumber, date: mkInfo.date ?? absentDate, subject, absentDate })
          makeupStudentInfo.set(sid, arr)
        }
      }
    }
  }

  // Helper: check if student has remaining makeup demand for a specific teacher and subject
  // targetDate restricts actual-result-based absences to only match slots AFTER the absence date
  // Always requires same teacher as original lesson
  const hasMakeupForTeacher = (studentId: string, teacherId: string, baseSubj: string, targetDate?: string): boolean => {
    const mkInfos = makeupStudentInfo.get(studentId)
    if (!mkInfos) return false
    return mkInfos.some((mk) => {
      if (mk.teacherId !== teacherId || mk.subject !== baseSubj) return false
      // For actual-result absences, only allow placement on dates strictly after the absence
      if (mk.absentDate && targetDate && targetDate <= mk.absentDate) return false
      return true
    })
  }

  // Pre-populate result with actual results (recorded slots) so student load counting includes them
  if (data.actualResults) {
    const seededSlots: string[] = []
    for (const [slot, results] of Object.entries(data.actualResults)) {
      const origSlot = data.assignments[slot] ?? []
      result[slot] = results.map((r) => {
        const orig = origSlot.find((a) => a.teacherId === r.teacherId)
        return {
          teacherId: r.teacherId,
          studentIds: [...r.studentIds],
          subject: r.subject,
          ...(r.studentSubjects ? { studentSubjects: { ...r.studentSubjects } } : {}),
          ...(orig?.isRegular ? { isRegular: true } : {}),
          ...(orig?.isGroupLesson ? { isGroupLesson: true } : {}),
          ...(orig?.regularMakeupInfo ? { regularMakeupInfo: { ...orig.regularMakeupInfo } } : {}),
        }
      })
      seededSlots.push(slot)
    }
    console.log('[AutoAssign] Seeded', seededSlots.length, 'recorded slots into result')
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

      // Skip assignments with no students (teacher-only)
      if (!assignment.isGroupLesson && assignment.studentIds.length === 0) {
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
        // All students removed — don't keep teacher-only assignment
        changeLog.push({ slot, action: '生徒全員解除', detail: `講師のみ残留破棄（予定変更のため）` })
      } else if (assignment.studentIds.length > 0) {
        cleaned.push(assignment)
      }
      // Skip assignments that already have no students (don't keep teacher-only)
    }

    if (cleaned.length > 0) {
      result[slot] = cleaned
    }
  }

  console.log('[AutoAssign] After Phase 1: result has', Object.keys(result).length, 'slots')

  // Consume makeup demand already satisfied by preserved assignments (actual results + Phase 1)
  // This prevents duplicate makeup assignments on re-assign
  for (const slotAssignments of Object.values(result)) {
    for (const assignment of slotAssignments) {
      if (!assignment.regularMakeupInfo) continue
      for (const [sid, mkInfo] of Object.entries(assignment.regularMakeupInfo)) {
        const mkInfos = makeupStudentInfo.get(sid)
        if (!mkInfos) continue
        // Match by dayOfWeek + slotNumber (the original absent slot), teacher may differ for actual-result absences
        const idx = mkInfos.findIndex(mk =>
          mk.dayOfWeek === mkInfo.dayOfWeek &&
          mk.slotNumber === mkInfo.slotNumber
        )
        if (idx >= 0) mkInfos.splice(idx, 1)
      }
    }
  }

  // Determine if this is the initial assignment or a re-run
  const isInitialAssignment = !Object.values(data.assignments).some(a => a && a.some(b => hasMeaningfulManualAssignment(b)))

  // Phase 1.5: Remove excess assignments when requested slots were reduced (ONLY on initial assignment)
  // On re-runs, preserve existing student assignments as-is
  if (isInitialAssignment) {
  // Include seeded recorded slot load so we correctly detect over-allocation
  const specialLoadMap = new Map<string, number>()
  for (const [, slotAssignments] of Object.entries(result)) {
    for (const assignment of slotAssignments) {
      if (assignment.isRegular) continue
      if (assignment.regularMakeupInfo) {
        for (const studentId of assignment.studentIds) {
          if (assignment.regularMakeupInfo[studentId]) continue
          const subj = getStudentSubject(assignment, studentId)
          const key = `${studentId}|${subj}`
          specialLoadMap.set(key, (specialLoadMap.get(key) ?? 0) + 1)
        }
      } else {
        for (const studentId of assignment.studentIds) {
          const subj = getStudentSubject(assignment, studentId)
          const key = `${studentId}|${subj}`
          specialLoadMap.set(key, (specialLoadMap.get(key) ?? 0) + 1)
        }
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
        const requested = (student?.subjectSlots ?? {})[subj] ?? 0
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
        } else {
          // All students removed by excess — remove this assignment entirely
          const slotArr = result[slot]
          if (slotArr) {
            const idx = slotArr.indexOf(assignment)
            if (idx >= 0) slotArr.splice(idx, 1)
          }
        }
      }
    }
  }
  } // end isInitialAssignment (Phase 1.5)

  // Phase 2: Fill empty student positions in existing assignments (including regular with empty spots)
  for (const slot of slots) {
    if (!result[slot] || result[slot].length === 0) continue
    const slotAssignments = result[slot]
    for (let idx = 0; idx < slotAssignments.length; idx++) {
      const assignment = slotAssignments[idx]
      if (assignment.isGroupLesson) continue
      if (assignment.studentIds.length >= 2) continue
      // On re-assign, preserve existing assignments with students to avoid unnecessary changes
      if (!isInitialAssignment && assignment.studentIds.length > 0) continue
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
        // Student must be able to learn at least one subject the teacher can teach (grade-aware), with remaining demand or makeup need
        const [slotDate] = slot.split('_')
        return student.subjects.some((baseSubj) => {
          if (!canTeachSubject(teacher.subjects, student.grade, baseSubj)) return false
          const requested = (student.subjectSlots ?? {})[baseSubj] ?? 0
          const allocated = countStudentSubjectLoad(result, student.id, baseSubj)
          return allocated < requested || hasMakeupForTeacher(student.id, teacher.id, baseSubj, slotDate)
        })
      })

      if (candidates.length > 0 && assignment.studentIds.length < 2) {
        const best = candidates.sort((a, b) => {
          const aRem = Object.values(a.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, a.id)
          const bRem = Object.values(b.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, b.id)
          return bRem - aRem
        })[0]
        // Pick the best viable subject for this new student
        const [bestSlotDate] = slot.split('_')
        const bestSubj = teachableBaseSubjects(teacher.subjects, best.grade).find((baseSubj) => {
          if (!best.subjects.includes(baseSubj)) return false
          const requested = (best.subjectSlots ?? {})[baseSubj] ?? 0
          const allocated = countStudentSubjectLoad(result, best.id, baseSubj)
          return allocated < requested || hasMakeupForTeacher(best.id, teacher.id, baseSubj, bestSlotDate)
        }) ?? assignment.subject
        // Reconstruct studentSubjects
        const studentSubjects: Record<string, string> = {}
        for (const sid of assignment.studentIds) {
          studentSubjects[sid] = getStudentSubject(assignment, sid)
        }
        studentSubjects[best.id] = bestSubj
        assignment.studentIds = [...assignment.studentIds, best.id]
        assignment.studentSubjects = studentSubjects
        // Check if this is a makeup student being added
        const mkInfos = makeupStudentInfo.get(best.id)
        const [mkDate2] = slot.split('_')
        const mkMatch = mkInfos?.find(mk => mk.teacherId === teacher.id && mk.subject === bestSubj && (!mk.absentDate || mkDate2 > mk.absentDate))
        if (mkMatch) {
          // Set regularMakeupInfo so ★ badge appears
          assignment.regularMakeupInfo = { ...(assignment.regularMakeupInfo ?? {}), [best.id]: { dayOfWeek: mkMatch.dayOfWeek, slotNumber: mkMatch.slotNumber, date: mkMatch.date } }
          const mkIdx = mkInfos!.indexOf(mkMatch)
          if (mkIdx >= 0) mkInfos!.splice(mkIdx, 1)
          const [mkDate] = slot.split('_')
          const mkSlotNum = getSlotNumber(slot)
          const fmtDate = (d: string) => { const [, m, day] = d.split('-'); return `${Number(m)}/${Number(day)}` }
          const origDate = mkMatch.date ? fmtDate(mkMatch.date) : ''
          const makeupDetail = `生徒希望入力で振替希望があったため自動振替（${best.name}: ${origDate} ${mkMatch.slotNumber}限 → ${fmtDate(mkDate)} ${mkSlotNum}限）`
          markMakeupPair(slot, assignment, makeupDetail)
        } else {
          const addChangeInfo = describeStudentSubmissionChange(best.id)
          const addReason = addChangeInfo || `${best.name}の${bestSubj}が未充足(残${(best.subjectSlots[bestSubj] ?? 0) - countStudentSubjectLoad(result, best.id, bestSubj)}コマ)のため空き枠に追加`
          markChangedPair(slot, assignment, `生徒追加: ${best.name}(${bestSubj})\n理由: ${addReason}`)
        }
        changeLog.push({ slot, action: '生徒追加', detail: `${best.name}(${bestSubj}) を追加` })
      }
    }
  }

  // Log demand summary before Phase 3
  {
    const demandSummary = data.students.map((s) => {
      const rem = Object.entries(s.subjectSlots)
        .map(([subj, req]) => ({ subj, rem: req - countStudentSubjectLoad(result, s.id, subj) }))
        .filter((x) => x.rem > 0)
      return rem.length > 0 ? `${s.name}: ${rem.map((x) => `${x.subj}残${x.rem}`).join(',')}` : null
    }).filter(Boolean)
    console.log('[AutoAssign] Before Phase 3: unmet demand:', demandSummary.length > 0 ? demandSummary.join('; ') : '(none)')
  }

  // Phase 3: Fill empty slots — minimize total teacher attendance dates, distribute students evenly

  const deskCountLimit = data.settings.deskCount ?? 0

  let slotCounter = 0
  for (const slot of slots) {
    // Yield every 5 slots to keep the page responsive
    if (++slotCounter % 5 === 0) {
      await yieldToMain()
      onProgress?.(slotCounter / slots.length)
    }

    const currentDate = slot.split('_')[0]
    const currentSlotNum = getSlotNumber(slot)

    // Skip slot 0 (午前) — only group lessons go there (auto-filled separately)
    if (currentSlotNum === 0) continue

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

    // Determine which students have a group lesson on this date (for groupContinuous card)
    const currentDayOfWeek = getIsoDayOfWeek(currentDate)
    const groupLessonsOnDate = (data.groupLessons ?? []).filter((gl) => gl.dayOfWeek === currentDayOfWeek)

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

      const teacherSlotsOnDate = getTeacherSlotNumbersOnDate(result, teacher.id, currentDate)
      const teacherIsConsecutive = teacherSlotsOnDate.some((n) => Math.abs(n - currentSlotNum) === 1)

      // Students that the teacher taught in the immediately previous slot (avoid same student consecutive)
      const prevSlotStudentIds = new Set(getTeacherPrevSlotStudentIds(result, teacher.id, currentDate, currentSlotNum))

      let bestPlan: { score: number; assignment: Assignment } | null = null

      for (const combo of [...candidates.map((s) => [s]), ...candidates.flatMap((l, i) => candidates.slice(i + 1).map((r) => [l, r]))]) {

        // Avoid assigning the same student to this teacher's consecutive slot
        const hasSameStudentConsecutive = combo.some((st) => prevSlotStudentIds.has(st.id))
        if (hasSameStudentConsecutive) continue

        // Check student-student constraints within this combo
        if (combo.length === 2 && constraintFor(data.constraints, combo[0].id, combo[1].id) === 'incompatible') continue

        // ── Hard filters ──
        // 一日上限2コマ (shared rule) — block if any student already has 2 slots today
        const anyOver2 = combo.some((st) => {
          const slotsOnDate = countStudentSlotsOnDate(result, st.id, currentDate)
          return slotsOnDate >= 2
        })
        if (anyOver2) continue

        // ── Constraint card hard filters ──
        let cardBlocked = false
        let totalCardScore = 0
        for (const st of combo) {
          const evalResult = evaluateConstraintCards(
            st, slot, result, data.settings.slotsPerDay,
            data.regularLessons, groupLessonsOnDate, teacher.id,
          )
          if (evalResult.blocked) { cardBlocked = true; break }
          totalCardScore += evalResult.score
        }
        if (cardBlocked) continue

        // --- Determine subject assignment (same or mixed) ---
        // Find base subjects ALL students in combo can learn and teacher can teach to ALL of them
        const commonBaseSubjects = (BASE_SUBJECTS as readonly string[]).filter((baseSubj) =>
          combo.every((student) => student.subjects.includes(baseSubj) && canTeachSubject(teacher.subjects, student.grade, baseSubj)),
        )
        const viableCommonSubjects = commonBaseSubjects.filter((baseSubj) =>
          combo.every((student) => {
            const requested = (student.subjectSlots ?? {})[baseSubj] ?? 0
            const allocated = countStudentSubjectLoad(result, student.id, baseSubj)
            return allocated < requested || hasMakeupForTeacher(student.id, teacher.id, baseSubj, currentDate)
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
            const req = (s1.subjectSlots ?? {})[baseSubj] ?? 0
            const alloc = countStudentSubjectLoad(result, s1.id, baseSubj)
            return alloc < req || hasMakeupForTeacher(s1.id, teacher.id, baseSubj, currentDate)
          })
          const s2Viable = teachableBaseSubjects(teacher.subjects, s2.grade).filter((baseSubj) => {
            if (!s2.subjects.includes(baseSubj)) return false
            const req = (s2.subjectSlots ?? {})[baseSubj] ?? 0
            const alloc = countStudentSubjectLoad(result, s2.id, baseSubj)
            return alloc < req || hasMakeupForTeacher(s2.id, teacher.id, baseSubj, currentDate)
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
        // ── Shared rule scoring (priority order) ──

        // 1. 2人ペアボーナス → 講師稼働率の最大化
        const pairBonus = combo.length === 2 ? 1000 : 0

        // 2. 既出勤日に追加 → 講師の出勤日数を最小化
        // NOTE: Teacher pre-sort already strongly prefers existing dates,
        // so scoring delta is kept moderate to avoid front-loading early dates.
        const attendanceBonus = isExistingDate ? 80 : -30

        // 3. 残コマ多数優先 → 生徒競合の場合残コマ数が多い生徒を優先
        let remainingSlotScore = 0
        for (const st of combo) {
          const totalRequested = Object.values(st.subjectSlots).reduce((s, c) => s + c, 0)
          const totalAssigned = countStudentLoad(result, st.id)
          remainingSlotScore += (totalRequested - totalAssigned) * 20
        }

        // ── Additional scoring factors (lower priority) ──

        // Teacher consecutive slot bonus
        const teacherConsecBonus = teacherIsConsecutive ? 50 : 0

        // Mixed-subject penalty: same subject pairs are slightly preferred
        const mixedSubjectPenalty = plan.isMixed ? -15 : 0

        // Consecutive same-subject penalty: avoid same student having the same subject in adjacent slots
        let consecutiveSameSubjectPenalty = 0
        for (const st of combo) {
          const subj = plan.isMixed ? (plan.studentSubjects[st.id] ?? '') : plan.subject
          const adjacentSubjects = getStudentSubjectsOnAdjacentSlots(result, st.id, currentDate, currentSlotNum)
          if (adjacentSubjects.includes(subj)) {
            // Cards that control multi-slot placement (twoConsecutive/twoWithGap/regularLink) want different subjects
            const cards = st.constraintCards ?? getDefaultConstraintCards(st.grade)
            const hasMultiSlotCard = cards.some((c) => c === 'twoConsecutive' || c === 'twoWithGap' || c === 'regularLink')
            if (hasMultiSlotCard) {
              consecutiveSameSubjectPenalty -= 100 // Stronger penalty when card wants diff subjects
            } else {
              consecutiveSameSubjectPenalty -= 20
            }
          }
        }

        // Makeup bonus: strongly prefer assigning makeup students to their regular teacher
        let makeupBonus = 0
        for (const st of combo) {
          const mkInfos = makeupStudentInfo.get(st.id)
          if (mkInfos && mkInfos.length > 0 && mkInfos.some(mk => mk.teacherId === teacher.id)) {
            makeupBonus += 200
          }
        }

        // Submission order bonus: earlier submitters get priority (max +15)
        let submissionBonus = 0
        for (const st of combo) {
          const submissionRank = submissionOrderMap.get(st.id) ?? maxRank
          submissionBonus += Math.max(0, 15 - submissionRank * 2)
        }

        // Teacher load balancing
        const teacherLoadPenalty = teacherLoad * 2

        const score =
          totalCardScore +            // Constraint cards (highest priority)
          pairBonus +                  // 2人ペアボーナス
          attendanceBonus +            // 既出勤日追加
          remainingSlotScore +         // 残コマ多数優先
          teacherConsecBonus +         // 講師連続コマ
          mixedSubjectPenalty +        // 混合科目
          consecutiveSameSubjectPenalty + // 隣接同科目
          makeupBonus +                // 振替ボーナス
          submissionBonus -            // 提出順
          teacherLoadPenalty           // 講師負荷

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
        // Check if any student is a makeup student assigned to their teacher
        const makeupSids = a.studentIds.filter((sid) => {
          const mkInfos = makeupStudentInfo.get(sid)
          return mkInfos && mkInfos.length > 0 && mkInfos.some(mk => mk.teacherId === a.teacherId && (!mk.absentDate || currentDate > mk.absentDate))
        })
        if (makeupSids.length > 0) {
          // Set regularMakeupInfo so the ★ badge appears alongside 振替
          const regularMakeupInfo: Record<string, { dayOfWeek: number; slotNumber: number; date?: string }> = { ...(a.regularMakeupInfo ?? {}) }
          for (const sid of makeupSids) {
            const mkInfos = makeupStudentInfo.get(sid)!
            const matchIdx = mkInfos.findIndex(mk => mk.teacherId === a.teacherId && (!mk.absentDate || currentDate > mk.absentDate))
            if (matchIdx >= 0) {
              const match = mkInfos[matchIdx]
              regularMakeupInfo[sid] = { dayOfWeek: match.dayOfWeek, slotNumber: match.slotNumber, date: match.date }
              mkInfos.splice(matchIdx, 1) // consume this makeup entry
            }
          }
          a.regularMakeupInfo = regularMakeupInfo
          // Build makeup-specific detail: include original and new slot info with dates
          const [mkDate] = slot.split('_')
          const mkSlotNum = getSlotNumber(slot)
          const fmtDate = (d: string) => { const [, m, day] = d.split('-'); return `${Number(m)}/${Number(day)}` }
          const makeupDetails = makeupSids.map((sid) => {
            const sName = data.students.find((s) => s.id === sid)?.name ?? '?'
            const info = regularMakeupInfo[sid]
            if (!info) return `${sName}: 振替`
            const origDate = info.date ? fmtDate(info.date) : ''
            return `${sName}: ${origDate} ${info.slotNumber}限 → ${fmtDate(mkDate)} ${mkSlotNum}限`
          }).join(', ')
          const makeupDetail = `生徒希望入力で振替希望があったため自動振替（${makeupDetails}）`
          markMakeupPair(slot, a, makeupDetail)
        } else {
          markAddedPair(slot, a, fullDetail)
        }
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

  const makeupPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(makeupPairSigSetBySlot)) {
    if (signatureSet.size > 0) makeupPairSignatures[slot] = [...signatureSet]
  }

  // Cleanup: remove assignments with no valid students (empty teacher-only pairs or stale student IDs)
  for (const slot of Object.keys(result)) {
    const cleaned = result[slot].filter((a) => {
      if (a.isGroupLesson) return true
      // Check that at least one student ID actually exists
      return a.studentIds.some(sid => sid && studentIds.has(sid))
    })
    if (cleaned.length > 0) {
      result[slot] = cleaned
    } else {
      delete result[slot]
    }
  }

  console.log('[AutoAssign] Done: result has', Object.keys(result).length, 'slots,', changeLog.length, 'changes')

  return { assignments: result, changeLog, changedPairSignatures, addedPairSignatures, makeupPairSignatures, changeDetails: changeDetailsBySlot }
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
