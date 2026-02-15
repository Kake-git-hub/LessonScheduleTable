export type PersonType = 'teacher' | 'student'

export type ConstraintType = 'incompatible' | 'recommended'

export type SubjectSlotRequest = Record<string, number>

export type Teacher = {
  id: string
  name: string
  subjects: string[]
  memo: string
}

export type Student = {
  id: string
  name: string
  grade: string
  subjects: string[]
  subjectSlots: SubjectSlotRequest
  unavailableDates: string[]
  memo: string
  submittedAt: number
}

export type PairConstraint = {
  id: string
  teacherId: string
  studentId: string
  type: ConstraintType
}

export type Assignment = {
  teacherId: string
  studentIds: string[]
  subject: string
}

export type RegularLesson = {
  id: string
  teacherId: string
  studentIds: string[]
  subject: string
  dayOfWeek: number
  slotNumber: number
}

export type SessionSettings = {
  name: string
  adminPassword: string
  startDate: string
  endDate: string
  slotsPerDay: number
  holidays: string[]
}

export type SessionData = {
  settings: SessionSettings
  subjects: string[]
  teachers: Teacher[]
  students: Student[]
  constraints: PairConstraint[]
  availability: Record<string, string[]>
  assignments: Record<string, Assignment[]>
  regularLessons: RegularLesson[]
}
