import type { ActualResult, Assignment, RegularLesson, SessionData } from '../types'
import { hasAvailability } from './constraints'
import { canTeachSubject } from './subjects'

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
      effective[slot] = results.map((r) => ({
        teacherId: r.teacherId,
        studentIds: [...r.studentIds],
        subject: r.subject,
        studentSubjects: r.studentSubjects ? { ...r.studentSubjects } : undefined,
      }))
    }
  }
  return effective
}

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
  allAssignments(assignments).filter((a) => a.studentIds.includes(studentId) && !a.isRegular).length

/** Count how many SPECIAL (non-regular) slots a student is assigned for a specific subject */
export const countStudentSubjectLoad = (
  assignments: Record<string, Assignment[]>,
  studentId: string,
  subject: string,
): number =>
  allAssignments(assignments).filter(
    (a) => a.studentIds.includes(studentId) && getStudentSubject(a, studentId) === subject && !a.isRegular,
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
      if (assignment.isRegular) continue

      if (!assignment.teacherId) {
        shortages.push({ slot, detail: '講師未設定' })
        continue
      }

      const teacher = data.teachers.find((item) => item.id === assignment.teacherId)
      if (!teacher) {
        shortages.push({ slot, detail: `講師ID ${assignment.teacherId} が未登録` })
        continue
      }

      if (!hasAvailability(data.availability, 'teacher', teacher.id, slot)) {
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
