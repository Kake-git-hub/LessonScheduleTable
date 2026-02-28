import type { ConstraintType, PairConstraint, PersonType, RegularLesson, SessionData, Student } from '../types'
import { personKey } from './schedule'

/** Check if a pair constraint exists between two persons (order-independent). */
export const constraintFor = (
  constraints: PairConstraint[],
  idA: string,
  idB: string,
): ConstraintType | null => {
  const hit = constraints.find((item) =>
    (item.personAId === idA && item.personBId === idB) ||
    (item.personAId === idB && item.personBId === idA),
  )
  return hit?.type ?? null
}

export const hasAvailability = (
  availability: SessionData['availability'],
  type: PersonType,
  id: string,
  slotKeyValue: string,
): boolean => {
  const key = personKey(type, id)
  return (availability[key] ?? []).includes(slotKeyValue)
}

export const isStudentAvailable = (student: Student, slotKey: string): boolean => {
  if (!student.submittedAt) return false
  if ((student.unavailableSlots ?? []).includes(slotKey)) return false
  const [date] = slotKey.split('_')
  return !student.unavailableDates.includes(date)
}

/** For mendan sessions: check if parent has positive availability for a slot */
export const isParentAvailableForMendan = (
  availability: SessionData['availability'],
  studentId: string,
  slotKey: string,
): boolean => {
  const key = personKey('student', studentId)
  return (availability[key] ?? []).includes(slotKey)
}

/** Check if a teacher-student pair appears in regular lessons */
export const isRegularLessonPair = (regularLessons: RegularLesson[], teacherId: string, studentId: string): boolean =>
  regularLessons.some((rl) => rl.teacherId === teacherId && rl.studentIds.includes(studentId))
