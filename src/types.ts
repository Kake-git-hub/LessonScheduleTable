export type PersonType = 'teacher' | 'student' | 'manager'

export type ConstraintType = 'incompatible' | 'recommended'

export type SubjectSlotRequest = Record<string, number>

// ── Slot constraint cards ─────────────────────────────────
export type SlotConstraintType =
  | 'consecutive'           // N コマ連続
  | 'gap-then-consecutive'  // M コマ空けて N コマ連続

export type SlotConstraintPriority = 'must' | 'prefer'

export type SlotConstraint = {
  id: string
  type: SlotConstraintType
  /** Priority: 'must' = hard constraint (block), 'prefer' = soft (scoring bonus) */
  priority: SlotConstraintPriority
  params: {
    /** Number of consecutive slots (for 'consecutive' and 'gap-then-consecutive') */
    count?: number
    /** Whether consecutive slots should use different subjects */
    diffSubject?: boolean
    /** Number of gap slots before the consecutive block (for 'gap-then-consecutive') */
    gapSlots?: number
  }
}

export type Manager = {
  id: string
  name: string
  email: string
}

export type Teacher = {
  id: string
  name: string
  email: string
  subjects: string[]
  memo: string
}

export type Student = {
  id: string
  name: string
  email: string
  grade: string
  subjects: string[]
  subjectSlots: SubjectSlotRequest
  unavailableDates: string[]
  preferredSlots: string[]
  unavailableSlots: string[]
  memo: string
  submittedAt: number
  /** Per-student slot scheduling constraints (e.g. "2コマ連続", "1コマ空けて連続") */
  slotConstraints?: SlotConstraint[]
}

export type PairConstraintPersonType = 'teacher' | 'student'

export type PairConstraint = {
  id: string
  personAId: string
  personBId: string
  personAType: PairConstraintPersonType
  personBType: PairConstraintPersonType
  type: ConstraintType
}

export type GradeConstraint = {
  id: string
  teacherId: string
  grade: string
  type: ConstraintType
  subjects?: string[]
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

/** Actual result for a single pair in a slot (recorded after the lesson). */
export type ActualResult = {
  teacherId: string
  studentIds: string[]
  subject: string
  studentSubjects?: Record<string, string>
}

export type RegularLesson = {
  id: string
  teacherId: string
  studentIds: string[]
  subject: string
  /** Per-student subject overrides. Maps studentId → base subject. */
  studentSubjects?: Record<string, string>
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
  sessionType?: 'lecture' | 'mendan'
  mendanStartHour?: number
  confirmed?: boolean
  createdAt?: number
  updatedAt?: number
  lastAutoAssignedAt?: number
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
  managers: Manager[]
  teachers: Teacher[]
  students: Student[]
  constraints: PairConstraint[]
  gradeConstraints: GradeConstraint[]
  availability: Record<string, string[]>
  assignments: Record<string, Assignment[]>
  regularLessons: RegularLesson[]
  autoAssignHighlights?: { added?: Record<string, string[]>; changed?: Record<string, string[]>; changeDetails?: Record<string, Record<string, string>> }
  teacherSubmittedAt?: Record<string, number>
  shareTokens?: Record<string, string>
  submissionLog?: SubmissionLogEntry[]
  /** Per-slot actual results (recorded after lesson is done). Key = slotKey. */
  actualResults?: Record<string, ActualResult[]>
  /** Global hourly rates per tier (A/B/C/D) for salary calculation. */
  tierRates?: { A: number; B: number; C: number; D: number }
}

export type MasterData = {
  managers: Manager[]
  teachers: Teacher[]
  students: Student[]
  constraints: PairConstraint[]
  gradeConstraints: GradeConstraint[]
  regularLessons: RegularLesson[]
}
