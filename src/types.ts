export type PersonType = 'teacher' | 'student'

export type ConstraintType = 'incompatible' | 'recommended'

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
  memo: string
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
  assignments: Record<string, Assignment>
}
