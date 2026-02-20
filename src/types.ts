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
  preferredSlots: string[]
  unavailableSlots: string[]
  memo: string
  submittedAt: number
}

export type PairConstraint = {
  id: string
  teacherId: string
  studentId: string
  type: ConstraintType
}

export type GradeConstraint = {
  id: string
  teacherId: string
  grade: string
  type: ConstraintType
}

export type Assignment = {
  teacherId: string
  studentIds: string[]
  subject: string
  /** Per-student subject overrides. Maps studentId → subject.
   *  When set, each student may learn a different subject in the same pair. */
  studentSubjects?: Record<string, string>
  isRegular?: boolean
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
  deskCount?: number
  submissionStartDate?: string
  submissionEndDate?: string
  createdAt?: number
  updatedAt?: number
}

/** A snapshot of a student's submission (or change delta). */
export type SubmissionLogEntry = {
  personId: string
  personType: 'student' | 'teacher'
  submittedAt: number
  /** 'initial' for first submission, 'update' for subsequent changes */
  type: 'initial' | 'update'
  /** The data snapshot at time of submission (student) */
  subjects?: string[]
  subjectSlots?: SubjectSlotRequest
  unavailableSlots?: string[]
  /** Teacher availability snapshot */
  availability?: string[]
}

export type SessionData = {
  settings: SessionSettings
  subjects: string[]
  teachers: Teacher[]
  students: Student[]
  constraints: PairConstraint[]
  gradeConstraints: GradeConstraint[]
  availability: Record<string, string[]>
  assignments: Record<string, Assignment[]>
  regularLessons: RegularLesson[]
  autoAssignHighlights?: { added?: Record<string, string[]>; changed?: Record<string, string[]> }
  teacherSubmittedAt?: Record<string, number>
  shareTokens?: Record<string, string>
  submissionLog?: SubmissionLogEntry[]
}

export type MasterData = {
  teachers: Teacher[]
  students: Student[]
  constraints: PairConstraint[]
  gradeConstraints: GradeConstraint[]
  regularLessons: RegularLesson[]
}
