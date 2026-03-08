import type { ActualResult, Assignment, RegularLesson, SessionData } from '../types'
import { hasAvailability } from './constraints'
import { canTeachSubject } from './subjects'

const hasTeacherAvailabilityWithRegularFallback = (data: SessionData, teacherId: string, slot: string): boolean => {
  if (hasAvailability(data.availability, 'teacher', teacherId, slot)) return true
  if ((data.teacherSubmittedAt?.[teacherId] ?? 0) > 0) return false
  const [date] = slot.split('_')
  const dayOfWeek = getIsoDayOfWeek(date)
  const slotNumber = getSlotNumber(slot)
  return data.regularLessons.some((lesson) => lesson.teacherId === teacherId && lesson.dayOfWeek === dayOfWeek && lesson.slotNumber === slotNumber)
}

type AssignmentLike = Assignment | ActualResult

export const getSlotNumber = (slotKey: string): number => {
  const [, slot] = slotKey.split('_')
  return Number.parseInt(slot, 10)
}

export const getIsoDayOfWeek = (isoDate: string): number => {
  const [year, month, day] = isoDate.split('-').map(Number)
  return new Date(Date.UTC(year, month - 1, day)).getUTCDay()
}

export const getSlotDayOfWeek = (slotKey: string): number => {
  const [date] = slotKey.split('_')
  return getIsoDayOfWeek(date)
}

export const allAssignments = (assignments: Record<string, Assignment[]>): Assignment[] =>
  Object.values(assignments).flat()

const normalizeStudentIds = (studentIds: string[]): string[] => studentIds.filter((sid): sid is string => {
  if (typeof sid !== 'string') return false
  const trimmed = sid.trim()
  return !!trimmed
})

const normalizeStudentSubjects = (
  studentIds: string[],
  studentSubjects?: Record<string, string>,
): Record<string, string> | undefined => {
  if (!studentSubjects) return undefined
  const normalized = studentIds.reduce<Record<string, string>>((acc, sid) => {
    const subject = studentSubjects[sid]
    if (typeof subject === 'string' && subject.trim()) {
      acc[sid] = subject
    }
    return acc
  }, {})
  return Object.keys(normalized).length > 0 ? normalized : undefined
}

const normalizeInfoMap = <T>(
  studentIds: string[],
  infoMap?: Record<string, T>,
): Record<string, T> | undefined => {
  if (!infoMap) return undefined
  const normalized = studentIds.reduce<Record<string, T>>((acc, sid) => {
    if (sid in infoMap) {
      acc[sid] = infoMap[sid]
    }
    return acc
  }, {})
  return Object.keys(normalized).length > 0 ? normalized : undefined
}

export const normalizeAssignment = <T extends AssignmentLike>(
  assignment: T,
): T => {
  const studentIds = normalizeStudentIds(assignment.studentIds)
  const studentSubjects = normalizeStudentSubjects(studentIds, assignment.studentSubjects)
  const regularMakeupInfo = normalizeInfoMap(studentIds, assignment.regularMakeupInfo)
  const regularSubstituteInfo = normalizeInfoMap(studentIds, assignment.regularSubstituteInfo)
  const primarySubject = studentIds.length > 0
    ? (studentSubjects?.[studentIds[0]] ?? assignment.subject)
    : assignment.subject

  const normalized = {
    ...assignment,
    studentIds,
    subject: primarySubject,
  } as T & {
    studentSubjects?: Record<string, string>
    regularMakeupInfo?: Record<string, { dayOfWeek: number; slotNumber: number; date?: string }>
    regularSubstituteInfo?: Record<string, { regularTeacherId: string; dayOfWeek: number; slotNumber: number; date?: string }>
  }

  if (studentSubjects) normalized.studentSubjects = studentSubjects
  else delete normalized.studentSubjects

  if (regularMakeupInfo) normalized.regularMakeupInfo = regularMakeupInfo
  else delete normalized.regularMakeupInfo

  if (regularSubstituteInfo) normalized.regularSubstituteInfo = regularSubstituteInfo
  else delete normalized.regularSubstituteInfo

  return normalized
}

/**
 * Build "effective" assignments: for each slot, use actual results if recorded,
 * otherwise use planned assignments.
 */
export const buildEffectiveAssignments = (
  assignments: Record<string, Assignment[]>,
  actualResults?: Record<string, ActualResult[]>,
): Record<string, Assignment[]> => {
  const effective: Record<string, Assignment[]> = { ...assignments }
  if (actualResults) {
    for (const [slot, results] of Object.entries(actualResults)) {
      const originals = assignments[slot] ?? []
      effective[slot] = results.map((r) => {
        // Preserve isRegular/isGroupLesson/regularMakeupInfo from the original assignment
        const orig = originals.find((a) => a.teacherId === r.teacherId)
        return normalizeAssignment({
          teacherId: r.teacherId,
          studentIds: [...r.studentIds],
          subject: r.subject,
          ...(r.studentSubjects ? { studentSubjects: { ...r.studentSubjects } } : {}),
          ...(r.isRegular || orig?.isRegular ? { isRegular: true } : {}),
          ...(r.isGroupLesson || orig?.isGroupLesson ? { isGroupLesson: true } : {}),
          ...(orig?.regularMakeupInfo ? { regularMakeupInfo: { ...orig.regularMakeupInfo } } : {}),
          ...(r.regularMakeupInfo ? { regularMakeupInfo: { ...r.regularMakeupInfo } } : {}),
          ...(orig?.regularSubstituteInfo ? { regularSubstituteInfo: { ...orig.regularSubstituteInfo } } : {}),
          ...(r.regularSubstituteInfo ? { regularSubstituteInfo: { ...r.regularSubstituteInfo } } : {}),
        })
      })
    }
  }
  return effective
}

export const isStudentRegularSubstituteAssignment = (assignment: Assignment, studentId: string): boolean =>
  !!assignment.regularSubstituteInfo?.[studentId]

export const isStudentRegularMakeupAssignment = (assignment: Assignment, studentId: string): boolean =>
  !!assignment.regularMakeupInfo?.[studentId]

/** Get the subject for a specific student in an assignment (supports per-student subjects). */
export const getStudentSubject = (a: Assignment, studentId: string): string =>
  a.studentSubjects?.[studentId] ?? a.subject

export const countTeacherLoad = (assignments: Record<string, Assignment[]>, teacherId: string): number =>
  allAssignments(assignments).filter((a) => a.teacherId === teacherId).length

/** Collect unique dates a teacher is already assigned to */
export const getTeacherAssignedDates = (assignments: Record<string, Assignment[]>, teacherId: string): Set<string> => {
  const dates = new Set<string>()
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slotAssignments.some((a) => a.teacherId === teacherId)) {
      dates.add(slot.split('_')[0])
    }
  }
  return dates
}

/** Get the slot numbers a teacher is assigned on a specific date */
export const getTeacherSlotNumbersOnDate = (assignments: Record<string, Assignment[]>, teacherId: string, date: string): number[] => {
  const nums: number[] = []
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slot.startsWith(`${date}_`)) {
      if (slotAssignments.some((a) => a.teacherId === teacherId)) {
        nums.push(getSlotNumber(slot))
      }
    }
  }
  return nums.sort((a, b) => a - b)
}

/** Get student IDs that a teacher taught in the previous slot on the same date */
export const getTeacherPrevSlotStudentIds = (assignments: Record<string, Assignment[]>, teacherId: string, date: string, slotNum: number): string[] => {
  const prevSlotKey = `${date}_${slotNum - 1}`
  const prevAssignments = assignments[prevSlotKey] ?? []
  for (const a of prevAssignments) {
    if (a.teacherId === teacherId) return a.studentIds
  }
  return []
}

/** Count how many slots a student is assigned on a specific date (including regular) */
export const countStudentSlotsOnDate = (assignments: Record<string, Assignment[]>, studentId: string, date: string): number => {
  let count = 0
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slot.startsWith(`${date}_`)) {
      if (slotAssignments.some((a) => a.studentIds.includes(studentId))) count++
    }
  }
  return count
}

/** Get the slot numbers a student is assigned on a specific date */
export const getStudentSlotNumbersOnDate = (assignments: Record<string, Assignment[]>, studentId: string, date: string): number[] => {
  const nums: number[] = []
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slot.startsWith(`${date}_`)) {
      if (slotAssignments.some((a) => a.studentIds.includes(studentId))) {
        nums.push(getSlotNumber(slot))
      }
    }
  }
  return nums.sort((a, b) => a - b)
}

/** Get the slot numbers a student is assigned on a specific date, excluding group lessons. */
export const getStudentNonGroupSlotNumbersOnDate = (assignments: Record<string, Assignment[]>, studentId: string, date: string): number[] => {
  const nums: number[] = []
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slot.startsWith(`${date}_`)) {
      if (slotAssignments.some((a) => !a.isGroupLesson && a.studentIds.includes(studentId))) {
        nums.push(getSlotNumber(slot))
      }
    }
  }
  return nums.sort((a, b) => a - b)
}

/** Count unique dates a student is assigned to */
export const countStudentAssignedDates = (assignments: Record<string, Assignment[]>, studentId: string): number => {
  const dates = new Set<string>()
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slotAssignments.some((a) => a.studentIds.includes(studentId))) {
      dates.add(slot.split('_')[0])
    }
  }
  return dates.size
}

/** Count how many SPECIAL (non-regular) slots a student is assigned */
export const countStudentLoad = (assignments: Record<string, Assignment[]>, studentId: string): number =>
  allAssignments(assignments).filter((a) => a.studentIds.includes(studentId) && !a.isRegular && !a.regularMakeupInfo?.[studentId] && !a.regularSubstituteInfo?.[studentId]).length

/** Count how many SPECIAL (non-regular) slots a student is assigned for a specific subject */
export const countStudentSubjectLoad = (
  assignments: Record<string, Assignment[]>,
  studentId: string,
  subject: string,
): number =>
  allAssignments(assignments).filter(
    (a) => a.studentIds.includes(studentId) && getStudentSubject(a, studentId) === subject && !a.isRegular && !a.regularMakeupInfo?.[studentId] && !a.regularSubstituteInfo?.[studentId],
  ).length

export type TeacherShortageEntry = {
  slot: string
  detail: string
}

export const collectTeacherShortages = (
  data: SessionData,
  assignments: Record<string, Assignment[]>,
): TeacherShortageEntry[] => {
  const shortages: TeacherShortageEntry[] = []
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    for (const assignment of slotAssignments) {
      if (assignment.isGroupLesson) continue

      if (!assignment.teacherId) {
        shortages.push({ slot, detail: assignment.teacherUnassignedReason ?? '講師未割当' })
        continue
      }

      const teacher = data.teachers.find((item) => item.id === assignment.teacherId)
      if (!teacher) {
        shortages.push({ slot, detail: `講師ID ${assignment.teacherId} が未登録` })
        continue
      }

      if (!hasTeacherAvailabilityWithRegularFallback(data, teacher.id, slot)) {
        shortages.push({ slot, detail: `${teacher.name} が出席不可` })
        continue
      }

      // Check subject compatibility per student (grade-aware)
      for (const sid of assignment.studentIds) {
        const student = data.students.find((s) => s.id === sid)
        if (!student) continue
        const subj = assignment.studentSubjects?.[sid] ?? assignment.subject
        if (subj && !canTeachSubject(teacher.subjects, student.grade, subj)) {
          shortages.push({ slot, detail: `${teacher.name} の担当外科目(${subj}) — ${student.name}` })
        }
      }
    }
  }
  return shortages
}

export const assignmentSignature = (assignment: Assignment): string => {
  const sortedStudents = [...assignment.studentIds].sort()
  const subjectPart = assignment.studentSubjects
    ? sortedStudents.map((sid) => `${sid}:${assignment.studentSubjects![sid] ?? assignment.subject}`).join('+')
    : `${assignment.subject}|${sortedStudents.join('+')}`
  return `${assignment.teacherId}|${subjectPart}|${assignment.isRegular ? 'R' : 'N'}`
}

export const hasMeaningfulManualAssignment = (assignment: Assignment): boolean =>
  !assignment.isRegular && !!(assignment.teacherId || assignment.subject || assignment.studentIds.length > 0)

/** Get teacher-student pair assignments on a specific date (for consecutive slot grouping) */
export const getTeacherStudentSlotsOnDate = (
  assignments: Record<string, Assignment[]>,
  teacherId: string,
  studentId: string,
  date: string,
): number[] => {
  const nums: number[] = []
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slot.startsWith(`${date}_`)) {
      if (slotAssignments.some((a) => a.teacherId === teacherId && a.studentIds.includes(studentId))) {
        nums.push(getSlotNumber(slot))
      }
    }
  }
  return nums.sort((a, b) => a - b)
}

/** Get subjects a student is assigned on adjacent (consecutive) slots on the same date */
export const getStudentSubjectsOnAdjacentSlots = (
  assignments: Record<string, Assignment[]>,
  studentId: string,
  date: string,
  slotNum: number,
): string[] => {
  const subjects: string[] = []
  for (const delta of [-1, 1]) {
    const adjacentSlot = `${date}_${slotNum + delta}`
    const slotAssigns = assignments[adjacentSlot]
    if (!slotAssigns) continue
    for (const a of slotAssigns) {
      if (a.studentIds.includes(studentId)) {
        subjects.push(getStudentSubject(a, studentId))
      }
    }
  }
  return subjects
}

export const findRegularLessonsForSlot = (
  regularLessons: RegularLesson[],
  slotKey: string,
): RegularLesson[] => {
  const dayOfWeek = getSlotDayOfWeek(slotKey)
  const slotNumber = getSlotNumber(slotKey)
  return regularLessons.filter((lesson) => lesson.dayOfWeek === dayOfWeek && lesson.slotNumber === slotNumber)
}

/** Get dates in range excluding holidays */
export const getDatesInRange = (settings: SessionData['settings']): string[] => {
  if (!settings.startDate || !settings.endDate) {
    return []
  }

  const start = new Date(`${settings.startDate}T00:00:00`)
  const end = new Date(`${settings.endDate}T00:00:00`)
  const holidaySet = new Set(settings.holidays)
  const dates: string[] = []

  for (let cursor = new Date(start); cursor <= end; ) {
    const y = cursor.getFullYear()
    const m = String(cursor.getMonth() + 1).padStart(2, '0')
    const d = String(cursor.getDate()).padStart(2, '0')
    const iso = `${y}-${m}-${d}`
    if (!holidaySet.has(iso)) {
      dates.push(iso)
    }
    cursor = new Date(cursor.getTime() + 24 * 60 * 60 * 1000)
  }

  return dates
}

export const getRegularSubjectProgress = (
  data: Pick<SessionData, 'settings' | 'regularLessons' | 'groupLessons'>,
  effectiveAssignments: Record<string, Assignment[]>,
  studentId: string,
): { desiredBySubject: Record<string, number>; assignedBySubject: Record<string, number> } => {
  const groupSubjectSet = new Set(
    (data.groupLessons ?? [])
      .filter((lesson) => lesson.studentIds.includes(studentId))
      .map((lesson) => lesson.subject),
  )

  const occurrences: { date: string; slot: string; dayOfWeek: number; slotNumber: number; subject: string }[] = []
  const desiredBySubject: Record<string, number> = {}

  for (const date of getDatesInRange(data.settings)) {
    const dayOfWeek = getIsoDayOfWeek(date)
    for (const lesson of data.regularLessons) {
      if (lesson.dayOfWeek !== dayOfWeek || !lesson.studentIds.includes(studentId)) continue
      const subject = lesson.studentSubjects?.[studentId] ?? lesson.subject
      if (groupSubjectSet.has(subject)) continue
      desiredBySubject[subject] = (desiredBySubject[subject] ?? 0) + 1
      occurrences.push({ date, slot: `${date}_${lesson.slotNumber}`, dayOfWeek, slotNumber: lesson.slotNumber, subject })
    }
  }

  const regularLikeAssignments = Object.entries(effectiveAssignments).flatMap(([slot, slotAssignments]) =>
    slotAssignments
      .filter((assignment) => assignment.studentIds.includes(studentId) && !assignment.isGroupLesson)
      .map((assignment) => ({ slot, assignment, subject: getStudentSubject(assignment, studentId), used: false })),
  )

  const assignedBySubject: Record<string, number> = {}
  for (const occurrence of occurrences) {
    const matchIndex = regularLikeAssignments.findIndex((entry) => {
      if (entry.used || entry.subject !== occurrence.subject) return false
      if (entry.assignment.isRegular && entry.slot === occurrence.slot) return true

      const makeupInfo = entry.assignment.regularMakeupInfo?.[studentId]
      if (makeupInfo && makeupInfo.dayOfWeek === occurrence.dayOfWeek && makeupInfo.slotNumber === occurrence.slotNumber) {
        return !makeupInfo.date || makeupInfo.date === occurrence.date
      }

      const substituteInfo = entry.assignment.regularSubstituteInfo?.[studentId]
      if (substituteInfo && substituteInfo.dayOfWeek === occurrence.dayOfWeek && substituteInfo.slotNumber === occurrence.slotNumber) {
        return !substituteInfo.date || substituteInfo.date === occurrence.date
      }

      return false
    })

    if (matchIndex >= 0) {
      regularLikeAssignments[matchIndex].used = true
      assignedBySubject[occurrence.subject] = (assignedBySubject[occurrence.subject] ?? 0) + 1
    }
  }

  return { desiredBySubject, assignedBySubject }
}
