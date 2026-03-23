export type PersonType = 'teacher' | 'student' | 'manager'

export type ConstraintType = 'incompatible' | 'recommended'

export type SubjectSlotRequest = Record<string, number>

export type RegularLessonAttendanceStatus = 'absent' | 'completed'

// ── Constraint cards ─────────────────────────────────
export type ConstraintCardType =
  | 'lateSlotNonExam'       // 受験生以外の後半コマ優先 (default)
  | 'groupContinuous'       // 集団後2コマ連続 (default)
  | 'forceRegularTeacher'   // 通常講師強制
  | 'priorityAssign'        // 優先割振 (default for 中3)
  | 'twoConsecutive'        // 2コマ連続
  | 'twoWithGap'            // 2コマ連続(一コマ空け)
  | 'oneSlotOnly'           // 1コマ上限
  | 'twoSlotLimit'          // 2コマ上限 (default)
  | 'threeSlotLimit'        // 3コマ上限
  | 'regularLink'           // 通常授業連結
  | 'earlySlotPreference'   // 小学生は2限寄り優先 (default)
  | 'lateSlotPreference'    // 中高生は5限寄り優先 (default)
  | 'avoidSlot1'            // 1限回避

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
  regularOnly?: boolean
  subjects: string[]
  subjectSlots: SubjectSlotRequest
  unavailableDates: string[]
  preferredSlots: string[]
  unavailableSlots: string[]
  regularLessonStatuses?: Record<string, RegularLessonAttendanceStatus>
  memo: string
  submittedAt: number
  /** Per-student constraint cards (e.g. "twoConsecutive", "oneSlotOnly") */
  constraintCards?: ConstraintCardType[]
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

export type RegularMakeupInfo = {
  dayOfWeek: number
  slotNumber: number
  date?: string
  reasonKind?: 'actual-absence'
}

export type StudentAbsenceRecord = {
  slot: string
  date: string
  dayOfWeek: number
  slotNumber: number
  teacherId: string
  lessonCategory?: 'regular' | 'makeup' | 'lecture'
  subject: string
}

export type Assignment = {
  teacherId: string
  studentIds: string[]
  subject: string
  /** Reason shown when a previously assigned teacher became unavailable and the pair was left unassigned. */
  teacherUnassignedReason?: string
  /** Original teacher ID when a pair was unassigned because the teacher became unavailable. */
  teacherUnavailableOriginalId?: string
  /** Per-student subject overrides. Maps studentId → subject.
   *  When set, each student may learn a different subject in the same pair. */
  studentSubjects?: Record<string, string>
  isRegular?: boolean
  isGroupLesson?: boolean
  /** Per-student regular-lesson makeup info.
   *  Key: studentId → which regular lesson slot this student is making up.
   *  Present when a student was absent from their regular lesson and assigned elsewhere. */
  regularMakeupInfo?: Record<string, RegularMakeupInfo>
  /** Per-student substitute info for a regular lesson handled by a different teacher.
   *  Key: studentId → original regular teacher + source regular lesson slot. */
  regularSubstituteInfo?: Record<string, { regularTeacherId: string; dayOfWeek: number; slotNumber: number; date?: string }>
  /** Per-student manual regular/makeup mark set via the UI badge toggle.
   *  'regular' = count as 通常, 'makeup' = count as 振替 */
  manualRegularMark?: Record<string, 'regular' | 'makeup'>
}

/** Actual result for a single pair in a slot (recorded after the lesson). */
export type ActualResult = {
  teacherId: string
  studentIds: string[]
  subject: string
  studentSubjects?: Record<string, string>
  regularMakeupInfo?: Record<string, RegularMakeupInfo>
  regularSubstituteInfo?: Record<string, { regularTeacherId: string; dayOfWeek: number; slotNumber: number; date?: string }>
  isRegular?: boolean
  isGroupLesson?: boolean
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

/** Group lesson: one subject, one teacher, multiple students, weekly at a fixed time */
export type GroupLesson = {
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
  regularOnly?: boolean
  subjects?: string[]
  subjectSlots?: SubjectSlotRequest
  unavailableDates?: string[]
  preferredSlots?: string[]
  unavailableSlots?: string[]
  regularLessonStatuses?: Record<string, RegularLessonAttendanceStatus>
  /** Teacher availability snapshot */
  availability?: string[]
}

export type PdfComparisonBaseline = {
  savedAt: number
  assignments: Record<string, Assignment[]>
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
  groupLessons: GroupLesson[]
  autoAssignHighlights?: { added?: Record<string, string[]>; changed?: Record<string, string[]>; makeup?: Record<string, string[]>; changeDetails?: Record<string, Record<string, string>>; unplacedMakeup?: { studentId: string; teacherId: string; subject: string; reason: string }[] }
  teacherSubmittedAt?: Record<string, number>
  shareTokens?: Record<string, string>
  submissionLog?: SubmissionLogEntry[]
  /** Per-slot actual results (recorded after lesson is done). Key = slotKey. */
  actualResults?: Record<string, ActualResult[]>
  pdfComparisonBaseline?: PdfComparisonBaseline
  /** Global hourly rates per tier (A/B/C/D) for salary calculation. */
  tierRates?: { A: number; B: number; C: number; D: number }
  /** Slots manually modified by the user — protected from regular-lesson auto-fill. */
  protectedManualSlots?: string[]
  /** Per-student absence history recorded from slot adjustment actions. */
  absenceRecords?: Record<string, StudentAbsenceRecord[]>
}

export type MasterData = {
  managers: Manager[]
  teachers: Teacher[]
  students: Student[]
  constraints: PairConstraint[]
  gradeConstraints: GradeConstraint[]
  regularLessons: RegularLesson[]
  groupLessons: GroupLesson[]
}
