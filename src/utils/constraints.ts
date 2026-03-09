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

export const getStudentRegularLessonStatus = (student: Student, slotKey: string): 'attend' | 'absent' | 'completed' => {
  const override = student.regularLessonStatuses?.[slotKey]
  if (override === 'completed') return 'completed'
  if (override === 'absent') return 'absent'
  if ((student.unavailableSlots ?? []).includes(slotKey)) return 'absent'
  const [date] = slotKey.split('_')
  return student.unavailableDates.includes(date) ? 'absent' : 'attend'
}

export const isStudentAvailableForRegularLesson = (student: Student, slotKey: string): boolean => {
  return getStudentRegularLessonStatus(student, slotKey) === 'attend'
}

export const isStudentRegularLessonCompleted = (student: Student, slotKey: string): boolean =>
  getStudentRegularLessonStatus(student, slotKey) === 'completed'

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
