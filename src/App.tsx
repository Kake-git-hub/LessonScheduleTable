import { useEffect, useMemo, useRef, useState } from 'react'
import { Link, Route, Routes, useLocation, useNavigate, useParams } from 'react-router-dom'
import XLSX from 'xlsx-js-style'
import './App.css'
import { deleteSession, initAuth, loadMasterData, loadSession, saveAndVerify, saveMasterData, saveSession, watchMasterData, watchSession, watchSessionsList } from './firebase'
import type {
  Assignment,
  ConstraintType,
  GradeConstraint,
  Manager,
  MasterData,
  PairConstraint,
  PersonType,
  RegularLesson,
  SessionData,
  Student,
  SubmissionLogEntry,
  Teacher,
} from './types'
import { buildSlotKeys, formatShortDate, mendanTimeLabel, personKey, slotLabel } from './utils/schedule'
import { downloadEmailReceiptPdf, downloadSubmissionReceiptPdf, exportSchedulePdf } from './utils/pdf'

const APP_VERSION = '1.0.0'

const GRADE_OPTIONS = ['小1', '小2', '小3', '小4', '小5', '小6', '中1', '中2', '中3', '高1', '高2', '高3']

const FIXED_SUBJECTS = ['英', '数', '国', '理', '社', 'IT']

const createId = (): string => {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID().slice(0, 8)
  }
  return Math.random().toString(36).slice(2, 10)
}

const emptySession = (): SessionData => ({
  settings: {
    name: '夏期講習',
    adminPassword: 'admin1234',
    startDate: '',
    endDate: '',
    slotsPerDay: 5,
    holidays: [],
    deskCount: 0,
    submissionStartDate: '',
    submissionEndDate: '',
  },
  subjects: FIXED_SUBJECTS,
  managers: [],
  teachers: [],
  students: [],
  constraints: [],
  gradeConstraints: [],
  availability: {},
  assignments: {},
  regularLessons: [],
  shareTokens: {},
  submissionLog: [],
})

const createTemplateSession = (): SessionData => {
  const settings: SessionData['settings'] = {
    name: '夏期講習テンプレート',
    adminPassword: 'admin1234',
    startDate: '2026-07-21',
    endDate: '2026-07-23',
    slotsPerDay: 5,
    holidays: [],
    deskCount: 0,
    submissionStartDate: '',
    submissionEndDate: '',
  }

  const subjects = FIXED_SUBJECTS

  const teachers: Teacher[] = [
    { id: 't001', name: '田中講師', email: '', subjects: ['数', '英'], memo: '数学メイン' },
    { id: 't002', name: '佐藤講師', email: '', subjects: ['英', '数'], memo: '英語メイン' },
  ]

  const students: Student[] = [
    {
      id: 's001',
      name: '青木 太郎',
      email: '',
      grade: '中3',
      subjects: ['数', '英'],
      subjectSlots: { 数: 3, 英: 2 },
      unavailableDates: ['2026-07-23'],
      preferredSlots: [],
      unavailableSlots: [],
      memo: '受験対策',
      submittedAt: Date.now() - 4000,
    },
    {
      id: 's002',
      name: '伊藤 花',
      email: '',
      grade: '中2',
      subjects: ['英'],
      subjectSlots: { 英: 3 },
      unavailableDates: [],
      preferredSlots: [],
      unavailableSlots: [],
      memo: '',
      submittedAt: Date.now() - 3000,
    },
    {
      id: 's003',
      name: '上田 陽介',
      email: '',
      grade: '高1',
      subjects: ['数'],
      subjectSlots: { 数: 3 },
      unavailableDates: ['2026-07-22'],
      preferredSlots: [],
      unavailableSlots: [],
      memo: '',
      submittedAt: Date.now() - 2000,
    },
    {
      id: 's004',
      name: '岡本 美咲',
      email: '',
      grade: '高2',
      subjects: ['英', '数'],
      subjectSlots: { 英: 2, 数: 2 },
      unavailableDates: [],
      preferredSlots: [],
      unavailableSlots: [],
      memo: '',
      submittedAt: Date.now() - 1000,
    },
  ]

  const constraints: PairConstraint[] = [
    { id: 'c001', teacherId: 't001', studentId: 's002', type: 'incompatible' },
  ]

  const slotKeys = buildSlotKeys(settings)
  const availability: SessionData['availability'] = {
    [personKey('teacher', 't001')]: slotKeys.filter((slot) => /_(1|2|3)$/.test(slot)),
    [personKey('teacher', 't002')]: slotKeys.filter((slot) => /_(1|2|3)$/.test(slot)),
    [personKey('student', 's001')]: slotKeys.filter((slot) => /_(1|2)$/.test(slot)),
    [personKey('student', 's002')]: slotKeys.filter((slot) => /_(2|3)$/.test(slot)),
    [personKey('student', 's003')]: slotKeys.filter((slot) => /_(1|3)$/.test(slot)),
    [personKey('student', 's004')]: slotKeys.filter((slot) => /_(1|2|3)$/.test(slot)),
  }

  const regularLessons: RegularLesson[] = [
    {
      id: 'r001',
      teacherId: 't001',
      studentIds: ['s001'],
      subject: '数',
      dayOfWeek: 1,
      slotNumber: 1,
    },
  ]

  return {
    settings,
    subjects,
    managers: [],
    teachers,
    students,
    constraints,
    gradeConstraints: [],
    availability,
    assignments: {},
    regularLessons,
  }
}

const useSessionData = (sessionId: string) => {
  const [data, setData] = useState<SessionData | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    setLoading(true)
    setError(null)
    const unsub = watchSession(
      sessionId,
      (value) => {
        setData(value)
        setLoading(false)
      },
      (err) => {
        console.error('watchSession error:', err)
        setError(err.message)
        setLoading(false)
      },
    )
    return () => unsub()
  }, [sessionId])

  return { data, setData, loading, error }
}

const constraintFor = (
  constraints: PairConstraint[],
  teacherId: string,
  studentId: string,
): ConstraintType | null => {
  const hit = constraints.find((item) => item.teacherId === teacherId && item.studentId === studentId)
  return hit?.type ?? null
}

const gradeConstraintFor = (
  gradeConstraints: GradeConstraint[],
  teacherId: string,
  grade: string,
  subject?: string,
): ConstraintType | null => {
  if (!grade) return null
  const hit = gradeConstraints.find((item) => {
    if (item.teacherId !== teacherId || item.grade !== grade) return false
    // If constraint has subjects specified, only match when subject matches
    if (item.subjects && item.subjects.length > 0) {
      if (!subject) return false // no subject provided → subject-specific constraint doesn't block
      return item.subjects.includes(subject)
    }
    return true // no subjects specified → universal constraint
  })
  return hit?.type ?? null
}

const hasAvailability = (
  availability: SessionData['availability'],
  type: PersonType,
  id: string,
  slotKeyValue: string,
): boolean => {
  const key = personKey(type, id)
  return (availability[key] ?? []).includes(slotKeyValue)
}

const allAssignments = (assignments: Record<string, Assignment[]>): Assignment[] =>
  Object.values(assignments).flat()

/** Get the subject for a specific student in an assignment (supports per-student subjects). */
const getStudentSubject = (a: Assignment, studentId: string): string =>
  a.studentSubjects?.[studentId] ?? a.subject

const countTeacherLoad = (assignments: Record<string, Assignment[]>, teacherId: string): number =>
  allAssignments(assignments).filter((a) => a.teacherId === teacherId).length

/** Collect unique dates a teacher is already assigned to */
const getTeacherAssignedDates = (assignments: Record<string, Assignment[]>, teacherId: string): Set<string> => {
  const dates = new Set<string>()
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slotAssignments.some((a) => a.teacherId === teacherId)) {
      dates.add(slot.split('_')[0])
    }
  }
  return dates
}

/** Get the slot numbers a teacher is assigned on a specific date */
const getTeacherSlotNumbersOnDate = (assignments: Record<string, Assignment[]>, teacherId: string, date: string): number[] => {
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
const getTeacherPrevSlotStudentIds = (assignments: Record<string, Assignment[]>, teacherId: string, date: string, slotNum: number): string[] => {
  const prevSlotKey = `${date}_${slotNum - 1}`
  const prevAssignments = assignments[prevSlotKey] ?? []
  for (const a of prevAssignments) {
    if (a.teacherId === teacherId) return a.studentIds
  }
  return []
}

/** Count how many slots a student is assigned on a specific date (including regular) */
const countStudentSlotsOnDate = (assignments: Record<string, Assignment[]>, studentId: string, date: string): number => {
  let count = 0
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slot.startsWith(`${date}_`)) {
      if (slotAssignments.some((a) => a.studentIds.includes(studentId))) count++
    }
  }
  return count
}

/** Get the slot numbers a student is assigned on a specific date */
const getStudentSlotNumbersOnDate = (assignments: Record<string, Assignment[]>, studentId: string, date: string): number[] => {
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
const countStudentAssignedDates = (assignments: Record<string, Assignment[]>, studentId: string): number => {
  const dates = new Set<string>()
  for (const [slot, slotAssignments] of Object.entries(assignments)) {
    if (slotAssignments.some((a) => a.studentIds.includes(studentId))) {
      dates.add(slot.split('_')[0])
    }
  }
  return dates.size
}

/** Count how many SPECIAL (non-regular) slots a student is assigned */
const countStudentLoad = (assignments: Record<string, Assignment[]>, studentId: string): number =>
  allAssignments(assignments).filter((a) => a.studentIds.includes(studentId) && !a.isRegular).length

/** Count how many SPECIAL (non-regular) slots a student is assigned for a specific subject */
const countStudentSubjectLoad = (
  assignments: Record<string, Assignment[]>,
  studentId: string,
  subject: string,
): number =>
  allAssignments(assignments).filter(
    (a) => a.studentIds.includes(studentId) && getStudentSubject(a, studentId) === subject && !a.isRegular,
  ).length

const isStudentAvailable = (student: Student, slotKey: string): boolean => {
  // Unsubmitted students (submittedAt === 0) are treated as unavailable for all dates
  if (!student.submittedAt) return false
  // Per-slot unavailability (new model)
  if ((student.unavailableSlots ?? []).includes(slotKey)) return false
  // Legacy: per-date unavailability (fallback for old data)
  const [date] = slotKey.split('_')
  return !student.unavailableDates.includes(date)
}

/** For mendan sessions: check if parent has positive availability for a slot */
const isParentAvailableForMendan = (
  availability: SessionData['availability'],
  studentId: string,
  slotKey: string,
): boolean => {
  const key = personKey('student', studentId)
  return (availability[key] ?? []).includes(slotKey)
}

const getIsoDayOfWeek = (isoDate: string): number => {
  const [year, month, day] = isoDate.split('-').map(Number)
  return new Date(Date.UTC(year, month - 1, day)).getUTCDay()
}

const getSlotDayOfWeek = (slotKey: string): number => {
  const [date] = slotKey.split('_')
  return getIsoDayOfWeek(date)
}

const getSlotNumber = (slotKey: string): number => {
  const [, slot] = slotKey.split('_')
  return Number.parseInt(slot, 10)
}

type TeacherShortageEntry = {
  slot: string
  detail: string
}

const collectTeacherShortages = (
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

      if (assignment.subject && !teacher.subjects.includes(assignment.subject)) {
        shortages.push({ slot, detail: `${teacher.name} の担当外科目(${assignment.subject})` })
      }
      // Check per-student subjects
      if (assignment.studentSubjects) {
        for (const [sid, subj] of Object.entries(assignment.studentSubjects)) {
          if (subj && !teacher.subjects.includes(subj)) {
            const sName = data.students.find((s) => s.id === sid)?.name ?? sid
            shortages.push({ slot, detail: `${teacher.name} の担当外科目(${subj}) — ${sName}` })
          }
        }
      }
    }
  }
  return shortages
}

const assignmentSignature = (assignment: Assignment): string => {
  const sortedStudents = [...assignment.studentIds].sort()
  const subjectPart = assignment.studentSubjects
    ? sortedStudents.map((sid) => `${sid}:${assignment.studentSubjects![sid] ?? assignment.subject}`).join('+')
    : `${assignment.subject}|${sortedStudents.join('+')}`
  return `${assignment.teacherId}|${subjectPart}|${assignment.isRegular ? 'R' : 'N'}`
}

const hasMeaningfulManualAssignment = (assignment: Assignment): boolean =>
  !assignment.isRegular && !!(assignment.teacherId || assignment.subject || assignment.studentIds.length > 0)

const ADMIN_PASSWORD_STORAGE_KEY = 'lst_admin_password_v1'
const readSavedAdminPassword = (): string => localStorage.getItem(ADMIN_PASSWORD_STORAGE_KEY) ?? 'admin1234'
const saveAdminPassword = (password: string): void => {
  localStorage.setItem(ADMIN_PASSWORD_STORAGE_KEY, password)
}

/** Check if a teacher-student pair appears in regular lessons */
const isRegularLessonPair = (regularLessons: RegularLesson[], teacherId: string, studentId: string): boolean =>
  regularLessons.some((rl) => rl.teacherId === teacherId && rl.studentIds.includes(studentId))

/** Get teacher-student pair assignments on a specific date (for consecutive slot grouping) */
const getTeacherStudentSlotsOnDate = (
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

const findRegularLessonsForSlot = (
  regularLessons: RegularLesson[],
  slotKey: string,
): RegularLesson[] => {
  const dayOfWeek = getSlotDayOfWeek(slotKey)
  const slotNumber = getSlotNumber(slotKey)
  return regularLessons.filter((lesson) => lesson.dayOfWeek === dayOfWeek && lesson.slotNumber === slotNumber)
}

// --- Incremental auto-assign: cleans up deleted people, fills gaps, keeps existing ---
interface ChangeLogEntry {
  slot: string
  action: string
  detail: string
}

const buildIncrementalAutoAssignments = (
  data: SessionData,
  slots: string[],
): { assignments: Record<string, Assignment[]>; changeLog: ChangeLogEntry[]; changedPairSignatures: Record<string, string[]>; addedPairSignatures: Record<string, string[]>; changeDetails: Record<string, Record<string, string>> } => {
  const changeLog: ChangeLogEntry[] = []
  const changedPairSigSetBySlot: Record<string, Set<string>> = {}
  const addedPairSigSetBySlot: Record<string, Set<string>> = {}
  const changeDetailsBySlot: Record<string, Record<string, string>> = {}

  // --- Helpers for detailed submission-based reasons ---
  const lastAutoAt = data.settings.lastAutoAssignedAt ?? 0

  /** Describe what changed in a student's submission since last auto-assign */
  const describeStudentSubmissionChange = (studentId: string): string => {
    const student = data.students.find((s) => s.id === studentId)
    if (!student) return ''
    // Find the most recent submission log entry for this student after lastAutoAt
    const recentEntries = (data.submissionLog ?? [])
      .filter((e) => e.personId === studentId && e.personType === 'student' && e.submittedAt > lastAutoAt)
      .sort((a, b) => b.submittedAt - a.submittedAt)
    if (recentEntries.length === 0) return ''

    const latest = recentEntries[0]
    const timeStr = new Date(latest.submittedAt).toLocaleDateString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    if (latest.type === 'initial') {
      // First submission after last auto-assign
      const slotsDetail = latest.subjectSlots
        ? Object.entries(latest.subjectSlots).map(([s, c]) => `${s}${c}コマ`).join(', ')
        : ''
      return `${student.name}: 前回コマ割り後に希望を新規提出(${timeStr})${slotsDetail ? ` [${slotsDetail}]` : ''}`
    }
    // Update submission — find the previous entry to show diff
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
  const teacherIds = new Set(data.teachers.map((t) => t.id))
  const studentIds = new Set(data.students.map((s) => s.id))
  const result: Record<string, Assignment[]> = {}

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

      // Collect all subjects needed in this assignment (per-student)
      const allNeededSubjects = assignment.studentSubjects
        ? [...new Set(Object.values(assignment.studentSubjects))]
        : assignment.subject ? [assignment.subject] : []

      // Check teacher still exists
      if (!teacherIds.has(assignment.teacherId)) {
        const usedTeachers = new Set(cleaned.map((a) => a.teacherId))
        const replacement = data.teachers.find((t) => {
          if (usedTeachers.has(t.id)) return false
          if (!hasAvailability(data.availability, 'teacher', t.id, slot)) return false
          return allNeededSubjects.every((subj) => t.subjects.includes(subj))
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
        const replacement = data.teachers.find((t) => {
          if (usedTeachers.has(t.id)) return false
          if (!hasAvailability(data.availability, 'teacher', t.id, slot)) return false
          return allNeededSubjects.every((subj) => t.subjects.includes(subj))
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
        const changedAssignment = { ...assignment, studentIds: [] }
        cleaned.push(changedAssignment)
        changeLog.push({ slot, action: '生徒全員解除', detail: `講師のみ残留（予定変更のため）` })
      } else {
        cleaned.push(assignment)
      }
    }

    if (cleaned.length > 0) {
      result[slot] = cleaned
    }
  }

  // Phase 2: Fill empty student positions in existing non-regular assignments
  // Phase 1.5: Remove excess assignments when requested slots were reduced
  const specialLoadMap = new Map<string, number>()
  for (const slot of slots) {
    const slotAssignments = result[slot] ?? []
    for (const assignment of slotAssignments) {
      if (assignment.isRegular) continue
      for (const studentId of assignment.studentIds) {
        const subj = getStudentSubject(assignment, studentId)
        const key = `${studentId}|${subj}`
        specialLoadMap.set(key, (specialLoadMap.get(key) ?? 0) + 1)
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
        const requested = student?.subjectSlots[subj] ?? 0
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
        }
      }
    }
  }

  // Phase 2: Fill empty student positions in existing non-regular assignments
  for (const slot of slots) {
    if (!result[slot] || result[slot].length === 0) continue
    const slotAssignments = result[slot]
    for (let idx = 0; idx < slotAssignments.length; idx++) {
      const assignment = slotAssignments[idx]
      if (assignment.isRegular) continue
      if (assignment.studentIds.length >= 2) continue
      if (!assignment.teacherId) continue

      const teacher = data.teachers.find((t) => t.id === assignment.teacherId)
      if (!teacher) continue

      const usedStudentIdsInSlot = new Set(slotAssignments.flatMap((a) => a.studentIds))

      const candidates = data.students.filter((student) => {
        if (usedStudentIdsInSlot.has(student.id)) return false
        if (!isStudentAvailable(student, slot)) return false
        if (constraintFor(data.constraints, teacher.id, student.id) === 'incompatible') return false
        if (gradeConstraintFor(data.gradeConstraints ?? [], teacher.id, student.grade) === 'incompatible') return false
        // Student must be able to learn at least one subject the teacher can teach, with remaining demand
        return teacher.subjects.some((subj) => {
          if (!student.subjects.includes(subj)) return false
          if (gradeConstraintFor(data.gradeConstraints ?? [], teacher.id, student.grade, subj) === 'incompatible') return false
          const requested = student.subjectSlots[subj] ?? 0
          const allocated = countStudentSubjectLoad(result, student.id, subj)
          return allocated < requested
        })
      })

      if (candidates.length > 0 && assignment.studentIds.length < 2) {
        const best = candidates.sort((a, b) => {
          const aRem = Object.values(a.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, a.id)
          const bRem = Object.values(b.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, b.id)
          return bRem - aRem
        })[0]
        // Pick the best viable subject for this new student
        const bestSubj = teacher.subjects.find((subj) => {
          if (!best.subjects.includes(subj)) return false
          const requested = best.subjectSlots[subj] ?? 0
          const allocated = countStudentSubjectLoad(result, best.id, subj)
          return allocated < requested
        }) ?? assignment.subject
        // Reconstruct studentSubjects
        const studentSubjects: Record<string, string> = {}
        for (const sid of assignment.studentIds) {
          studentSubjects[sid] = getStudentSubject(assignment, sid)
        }
        studentSubjects[best.id] = bestSubj
        assignment.studentIds = [...assignment.studentIds, best.id]
        assignment.studentSubjects = studentSubjects
        const addChangeInfo = describeStudentSubmissionChange(best.id)
        const addReason = addChangeInfo || `${best.name}の${bestSubj}が未充足(残${(best.subjectSlots[bestSubj] ?? 0) - countStudentSubjectLoad(result, best.id, bestSubj)}コマ)のため空き枠に追加`
        markChangedPair(slot, assignment, `生徒追加: ${best.name}(${bestSubj})\n理由: ${addReason}`)
        changeLog.push({ slot, action: '生徒追加', detail: `${best.name}(${bestSubj}) を追加` })
      }
    }
  }

  // Phase 3: Fill empty slots — minimize total teacher attendance dates, distribute students evenly
  // Process slots date-by-date, within each date prefer using teachers already assigned that day

  // Compute date ordering for first-half bias
  const allDatesInOrder: string[] = []
  for (const s of slots) {
    const d = s.split('_')[0]
    if (allDatesInOrder.length === 0 || allDatesInOrder[allDatesInOrder.length - 1] !== d) {
      allDatesInOrder.push(d)
    }
  }
  const totalDates = allDatesInOrder.length
  const dateIndexMap = new Map<string, number>()
  for (let i = 0; i < totalDates; i++) dateIndexMap.set(allDatesInOrder[i], i)

  const deskCountLimit = data.settings.deskCount ?? 0

  for (const slot of slots) {
    const currentDate = slot.split('_')[0]
    const currentSlotNum = getSlotNumber(slot)

    // Initialize from existing assignments (allow adding more pairs to a slot)
    const existingAssignments = result[slot] ?? []
    const slotAssignments: Assignment[] = [...existingAssignments]
    const usedTeacherIdsInSlot = new Set<string>(existingAssignments.map((a) => a.teacherId))
    const usedStudentIdsInSlot = new Set<string>(existingAssignments.flatMap((a) => a.studentIds))
    // Skip slots where all assignments are regular lessons (protected)
    if (existingAssignments.length > 0 && existingAssignments.every((a) => a.isRegular)) {
      continue
    }
    // Skip slots at desk count limit
    if (deskCountLimit > 0 && slotAssignments.length >= deskCountLimit) {
      continue
    }

    // First-half bias: earlier dates get a bonus (max +25 for first date, 0 for last)
    const dateIdx = dateIndexMap.get(currentDate) ?? 0
    const firstHalfBonus = totalDates > 1 ? Math.round(25 * (1 - dateIdx / (totalDates - 1))) : 0

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
        if (gradeConstraintFor(data.gradeConstraints ?? [], teacher.id, student.grade) === 'incompatible') return false
        return teacher.subjects.some((subject) => {
          if (!student.subjects.includes(subject)) return false
          if (gradeConstraintFor(data.gradeConstraints ?? [], teacher.id, student.grade, subject) === 'incompatible') return false
          return true
        })
      })

      if (candidates.length === 0) continue

      const teacherDates = getTeacherAssignedDates(result, teacher.id)
      const isExistingDate = teacherDates.has(currentDate)
      const teacherLoad = countTeacherLoad(result, teacher.id)

      // Teacher consecutive slot bonus
      const teacherSlotsOnDate = getTeacherSlotNumbersOnDate(result, teacher.id, currentDate)
      const teacherIsConsecutive = teacherSlotsOnDate.some((n) => Math.abs(n - currentSlotNum) === 1)
      const teacherConsecutiveBonus = teacherIsConsecutive ? 20 : 0

      // Students that the teacher taught in the immediately previous slot (avoid same student consecutive)
      const prevSlotStudentIds = new Set(getTeacherPrevSlotStudentIds(result, teacher.id, currentDate, currentSlotNum))

      let bestPlan: { score: number; assignment: Assignment } | null = null

      for (const combo of [...candidates.map((s) => [s]), ...candidates.flatMap((l, i) => candidates.slice(i + 1).map((r) => [l, r]))]) {
        // Avoid assigning the same student to this teacher's consecutive slot
        const hasSameStudentConsecutive = combo.some((st) => prevSlotStudentIds.has(st.id))
        if (hasSameStudentConsecutive) continue

        // --- Determine subject assignment (same or mixed) ---
        // Try same-subject first (shared by all students) — preferred
        const commonSubjects = teacher.subjects.filter((subject) =>
          combo.every((student) => student.subjects.includes(subject)),
        )
        const viableCommonSubjects = commonSubjects.filter((subject) =>
          combo.every((student) => {
            const requested = student.subjectSlots[subject] ?? 0
            const allocated = countStudentSubjectLoad(result, student.id, subject)
            return allocated < requested
          }),
        )

        // For 2-student combos: also try mixed-subject pairing (each student gets their own subject)
        type SubjectPlan = { isMixed: false; subject: string } | { isMixed: true; studentSubjects: Record<string, string>; primarySubject: string }
        const subjectPlans: SubjectPlan[] = []

        // Add same-subject plans
        for (const subj of viableCommonSubjects) {
          subjectPlans.push({ isMixed: false, subject: subj })
        }

        // Add mixed-subject plans for 2-student combos
        if (combo.length === 2) {
          const [s1, s2] = combo
          const s1Viable = teacher.subjects.filter((subj) => {
            if (!s1.subjects.includes(subj)) return false
            const req = s1.subjectSlots[subj] ?? 0
            const alloc = countStudentSubjectLoad(result, s1.id, subj)
            return alloc < req
          })
          const s2Viable = teacher.subjects.filter((subj) => {
            if (!s2.subjects.includes(subj)) return false
            const req = s2.subjectSlots[subj] ?? 0
            const alloc = countStudentSubjectLoad(result, s2.id, subj)
            return alloc < req
          })
          // Only add mixed plans where subjects actually differ
          for (const subj1 of s1Viable) {
            for (const subj2 of s2Viable) {
              if (subj1 === subj2) continue // same-subject already covered above
              subjectPlans.push({
                isMixed: true,
                studentSubjects: { [s1.id]: subj1, [s2.id]: subj2 },
                primarySubject: subj1, // Use first student's subject as the display subject
              })
            }
          }
        }

        if (subjectPlans.length === 0) continue

        for (const plan of subjectPlans) {
        // --- Student distribution scoring ---
        let studentScore = 0
        for (const st of combo) {
          const slotsOnDate = countStudentSlotsOnDate(result, st.id, currentDate)
          const existingSlotNums = getStudentSlotNumbersOnDate(result, st.id, currentDate)

          // Penalty for same-day multiple slots (avoid if possible)
          if (slotsOnDate > 0) {
            studentScore -= 60
            // If forced, strongly reward consecutive slots / penalize non-consecutive
            const isConsecutive = existingSlotNums.some(
              (n) => Math.abs(n - currentSlotNum) === 1,
            )
            if (isConsecutive) {
              studentScore += 50
            } else {
              studentScore -= 30
            }
          }

          // Prefer students with more remaining slots (even distribution)
          const totalRequested = Object.values(st.subjectSlots).reduce((s, c) => s + c, 0)
          const totalAssigned = countStudentLoad(result, st.id)
          studentScore += (totalRequested - totalAssigned) * 10

          // Prefer students with fewer assigned dates (spread across days)
          const assignedDates = countStudentAssignedDates(result, st.id)
          studentScore -= assignedDates * 5

          // Submission order bonus: earlier submitters get priority (max +15)
          const submissionRank = submissionOrderMap.get(st.id) ?? maxRank
          studentScore += Math.max(0, 15 - submissionRank * 2)
        }

        // Regular lesson pair bonus: prefer assigning regular-lesson teacher-student combos
        const regularPairBonus = combo.reduce((s, st) =>
          s + (isRegularLessonPair(data.regularLessons, teacher.id, st.id) ? 30 : 0), 0)

        // Same-day same-pair consecutive bonus: if this teacher+student pair already
        // exists on this date, strongly prefer making it consecutive
        let pairConsecutiveBonus = 0
        for (const st of combo) {
          const existingPairSlots = getTeacherStudentSlotsOnDate(result, teacher.id, st.id, currentDate)
          if (existingPairSlots.length > 0) {
            const isConsecutive = existingPairSlots.some((n) => Math.abs(n - currentSlotNum) === 1)
            pairConsecutiveBonus += isConsecutive ? 60 : -40
          }
        }

        // Mixed-subject penalty: same subject pairs are slightly preferred
        const mixedSubjectPenalty = plan.isMixed ? -15 : 0

        const score = 100 +
          (isExistingDate ? 80 : -50) +  // Very strong preference for reusing existing dates
          teacherConsecutiveBonus +  // Teacher consecutive slot bonus
          firstHalfBonus +  // First-half bias (max +25)
          regularPairBonus +  // Regular lesson pair preference
          pairConsecutiveBonus +  // Same teacher+student consecutive on same day
          (combo.length === 2 ? 30 : 0) +  // 2-person pair bonus
          mixedSubjectPenalty +  // Slight penalty for mixed subjects
          studentScore -
          teacherLoad * 2

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
        markAddedPair(slot, a, fullDetail)
        changeLog.push({ slot, action: '新規割当', detail: fullDetail })
      }
    }
  }

  const changedPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(changedPairSigSetBySlot)) {
    // Exclude signatures that are in the added set (added takes priority)
    const addedSet = addedPairSigSetBySlot[slot] ?? new Set<string>()
    const filtered = [...signatureSet].filter((sig) => !addedSet.has(sig))
    if (filtered.length > 0) changedPairSignatures[slot] = filtered
  }

  const addedPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(addedPairSigSetBySlot)) {
    if (signatureSet.size > 0) addedPairSignatures[slot] = [...signatureSet]
  }

  return { assignments: result, changeLog, changedPairSignatures, addedPairSignatures, changeDetails: changeDetailsBySlot }
}

/** Mendan (interview) FCFS auto-assign: each parent gets exactly 1 slot with 1 manager */
const buildMendanAutoAssignments = (
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

const emptyMasterData = (): MasterData => ({
  managers: [],
  teachers: [],
  students: [],
  constraints: [],
  gradeConstraints: [],
  regularLessons: [],
})

const HomePage = () => {
  const navigate = useNavigate()
  const location = useLocation()
  const [unlocked, setUnlocked] = useState(false)
  const [adminPassword, setAdminPassword] = useState(readSavedAdminPassword())
  const [sessions, setSessions] = useState<{ id: string; name: string; createdAt: number; updatedAt: number }[]>([])
  const [masterData, setMasterData] = useState<MasterData | null>(null)
  const [newYear, setNewYear] = useState(String(new Date().getFullYear()))
  const [newTerm, setNewTerm] = useState<'spring' | 'summer' | 'winter' | 'spring-mendan' | 'summer-mendan' | 'winter-mendan'>('summer')
  const [newSessionId, setNewSessionId] = useState('')
  const [newSessionName, setNewSessionName] = useState('')
  const [newStartDate, setNewStartDate] = useState('')
  const [newEndDate, setNewEndDate] = useState('')
  const [newSubmissionStart, setNewSubmissionStart] = useState('')
  const [newSubmissionEnd, setNewSubmissionEnd] = useState('')
  const [newDeskCount, setNewDeskCount] = useState(0)
  const [newHolidays, setNewHolidays] = useState<string[]>([])

  // Master data form state
  const [managerName, setManagerName] = useState('')
  const [managerEmail, setManagerEmail] = useState('')
  const [teacherName, setTeacherName] = useState('')
  const [teacherEmail, setTeacherEmail] = useState('')
  const [teacherSubjects, setTeacherSubjects] = useState<string[]>([])
  const [teacherMemo, setTeacherMemo] = useState('')
  const [studentName, setStudentName] = useState('')
  const [studentEmail, setStudentEmail] = useState('')
  const [studentGrade, setStudentGrade] = useState('')
  const [constraintTeacherId, setConstraintTeacherId] = useState('')
  const [constraintStudentId, setConstraintStudentId] = useState('')
  const [constraintType, setConstraintType] = useState<ConstraintType>('incompatible')
  const [gradeConstraintTeacherId, setGradeConstraintTeacherId] = useState('')
  const [gradeConstraintGrade, setGradeConstraintGrade] = useState('')
  const [gradeConstraintType, setGradeConstraintType] = useState<ConstraintType>('incompatible')
  const [gradeConstraintSubjects, setGradeConstraintSubjects] = useState<string[]>([])
  const [regularTeacherId, setRegularTeacherId] = useState('')
  const [regularStudent1Id, setRegularStudent1Id] = useState('')
  const [regularStudent2Id, setRegularStudent2Id] = useState('')
  const [regularSubject, setRegularSubject] = useState('')
  const [regularDayOfWeek, setRegularDayOfWeek] = useState('')
  const [regularSlotNumber, setRegularSlotNumber] = useState('')
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Editing state for master data
  const [editingManagerId, setEditingManagerId] = useState<string | null>(null)
  const [editManagerName, setEditManagerName] = useState('')
  const [editManagerEmail, setEditManagerEmail] = useState('')
  const [editingTeacherId, setEditingTeacherId] = useState<string | null>(null)
  const [editTeacherName, setEditTeacherName] = useState('')
  const [editTeacherEmail, setEditTeacherEmail] = useState('')
  const [editTeacherSubjects, setEditTeacherSubjects] = useState<string[]>([])
  const [editTeacherMemo, setEditTeacherMemo] = useState('')
  const [editingStudentId, setEditingStudentId] = useState<string | null>(null)
  const [editStudentName, setEditStudentName] = useState('')
  const [editStudentEmail, setEditStudentEmail] = useState('')
  const [editStudentGrade, setEditStudentGrade] = useState('')

  useEffect(() => {
    if (import.meta.env.DEV) {
      navigate('/boot', { replace: true })
    }
  }, [navigate])

  useEffect(() => {
    const directHome = (location.state as { directHome?: boolean } | null)?.directHome === true
    if (!directHome) return
    setUnlocked(true)
  }, [location.state])

  useEffect(() => {
    const year = Number.parseInt(newYear, 10)
    const safeYear = Number.isNaN(year) ? new Date().getFullYear() : year
    const termLabels: Record<string, string> = {
      spring: '春期講習', summer: '夏期講習', winter: '冬期講習',
      'spring-mendan': '春期面談', 'summer-mendan': '夏期面談', 'winter-mendan': '冬期面談',
    }
    const label = termLabels[newTerm] ?? '夏期講習'
    const idTerm = newTerm
    setNewSessionId(`${safeYear}-${idTerm}`)
    setNewSessionName(`${safeYear} ${label}`)
  }, [newTerm, newYear])

  useEffect(() => {
    if (!unlocked) return
    const unsub1 = watchSessionsList((items) => setSessions(items))
    const unsub2 = watchMasterData((md) => {
      if (md) {
        setMasterData(md)
      } else {
        const empty = emptyMasterData()
        saveMasterData(empty).catch(() => {})
        setMasterData(empty)
      }
    })
    return () => { unsub1(); unsub2() }
  }, [unlocked])

  // --- Master data helpers ---
  const updateMaster = async (updater: (current: MasterData) => MasterData): Promise<void> => {
    if (!masterData) return
    const next = updater(masterData)
    setMasterData(next)
    await saveMasterData(next)
  }

  const addManager = async (): Promise<void> => {
    if (!managerName.trim()) return
    const manager: Manager = { id: createId(), name: managerName.trim(), email: managerEmail.trim() }
    await updateMaster((c) => ({ ...c, managers: [...(c.managers ?? []), manager] }))
    setManagerName(''); setManagerEmail('')
  }

  const addTeacher = async (): Promise<void> => {
    if (!teacherName.trim()) return
    const teacher: Teacher = { id: createId(), name: teacherName.trim(), email: teacherEmail.trim(), subjects: teacherSubjects, memo: teacherMemo.trim() }
    await updateMaster((c) => ({ ...c, teachers: [...c.teachers, teacher] }))
    setTeacherName(''); setTeacherEmail(''); setTeacherSubjects([]); setTeacherMemo('')
  }

  const addStudent = async (): Promise<void> => {
    if (!studentName.trim()) return
    const student: Student = {
      id: createId(), name: studentName.trim(), email: studentEmail.trim(), grade: studentGrade.trim(),
      subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo: '', submittedAt: 0,
    }
    await updateMaster((c) => ({ ...c, students: [...c.students, student] }))
    setStudentName(''); setStudentEmail(''); setStudentGrade('')
  }

  const upsertConstraint = async (): Promise<void> => {
    if (!constraintTeacherId || !constraintStudentId || !masterData) return
    const nc: PairConstraint = { id: createId(), teacherId: constraintTeacherId, studentId: constraintStudentId, type: constraintType }
    await updateMaster((c) => {
      const filtered = c.constraints.filter((i) => !(i.teacherId === constraintTeacherId && i.studentId === constraintStudentId))
      return { ...c, constraints: [...filtered, nc] }
    })
  }

  const upsertGradeConstraint = async (): Promise<void> => {
    if (!gradeConstraintTeacherId || !gradeConstraintGrade || !masterData) return
    const nc: GradeConstraint = {
      id: createId(),
      teacherId: gradeConstraintTeacherId,
      grade: gradeConstraintGrade,
      type: gradeConstraintType,
      ...(gradeConstraintSubjects.length > 0 ? { subjects: gradeConstraintSubjects } : {}),
    }
    await updateMaster((c) => {
      const filtered = (c.gradeConstraints ?? []).filter((i) => !(i.teacherId === gradeConstraintTeacherId && i.grade === gradeConstraintGrade))
      return { ...c, gradeConstraints: [...filtered, nc] }
    })
    setGradeConstraintSubjects([])
  }

  const addRegularLesson = async (): Promise<void> => {
    const studentIds = [regularStudent1Id, regularStudent2Id].filter(Boolean)
    if (!regularTeacherId || studentIds.length === 0 || !regularSubject || !regularDayOfWeek || !regularSlotNumber) return
    const nl: RegularLesson = {
      id: createId(), teacherId: regularTeacherId, studentIds, subject: regularSubject,
      dayOfWeek: Number.parseInt(regularDayOfWeek, 10), slotNumber: Number.parseInt(regularSlotNumber, 10),
    }
    await updateMaster((c) => ({ ...c, regularLessons: [...c.regularLessons, nl] }))
    setRegularTeacherId(''); setRegularStudent1Id(''); setRegularStudent2Id('')
    setRegularSubject(''); setRegularDayOfWeek(''); setRegularSlotNumber('')
  }

  const startEditManager = (m: Manager): void => {
    setEditingManagerId(m.id); setEditManagerName(m.name); setEditManagerEmail(m.email || '')
  }
  const saveEditManager = async (): Promise<void> => {
    if (!editingManagerId || !editManagerName.trim()) return
    await updateMaster((c) => ({ ...c, managers: (c.managers ?? []).map((m) => m.id === editingManagerId ? { ...m, name: editManagerName.trim(), email: editManagerEmail.trim() } : m) }))
    setEditingManagerId(null)
  }

  const startEditTeacher = (t: Teacher): void => {
    setEditingTeacherId(t.id); setEditTeacherName(t.name); setEditTeacherEmail(t.email || '')
    setEditTeacherSubjects([...t.subjects]); setEditTeacherMemo(t.memo || '')
  }
  const saveEditTeacher = async (): Promise<void> => {
    if (!editingTeacherId || !editTeacherName.trim()) return
    await updateMaster((c) => ({ ...c, teachers: c.teachers.map((t) => t.id === editingTeacherId ? { ...t, name: editTeacherName.trim(), email: editTeacherEmail.trim(), subjects: editTeacherSubjects, memo: editTeacherMemo.trim() } : t) }))
    setEditingTeacherId(null)
  }

  const startEditStudent = (s: Student): void => {
    setEditingStudentId(s.id); setEditStudentName(s.name); setEditStudentEmail(s.email || '')
    setEditStudentGrade(s.grade || '')
  }
  const saveEditStudent = async (): Promise<void> => {
    if (!editingStudentId || !editStudentName.trim()) return
    await updateMaster((c) => ({ ...c, students: c.students.map((s) => s.id === editingStudentId ? { ...s, name: editStudentName.trim(), email: editStudentEmail.trim(), grade: editStudentGrade.trim() } : s) }))
    setEditingStudentId(null)
  }

  const removeManager = async (managerId: string): Promise<void> => {
    if (!window.confirm('このマネージャーを削除しますか？')) return
    await updateMaster((c) => ({ ...c, managers: (c.managers ?? []).filter((m) => m.id !== managerId) }))
  }

  const removeTeacher = async (teacherId: string): Promise<void> => {
    if (!window.confirm('この講師を削除しますか？')) return
    await updateMaster((c) => ({ ...c, teachers: c.teachers.filter((t) => t.id !== teacherId) }))
  }

  const removeStudent = async (studentId: string): Promise<void> => {
    if (!window.confirm('この生徒を削除しますか？')) return
    await updateMaster((c) => ({ ...c, students: c.students.filter((s) => s.id !== studentId) }))
  }

  const removeConstraint = async (constraintId: string): Promise<void> => {
    await updateMaster((c) => ({ ...c, constraints: c.constraints.filter((x) => x.id !== constraintId) }))
  }

  const removeGradeConstraint = async (constraintId: string): Promise<void> => {
    await updateMaster((c) => ({ ...c, gradeConstraints: (c.gradeConstraints ?? []).filter((x) => x.id !== constraintId) }))
  }

  const removeRegularLesson = async (lessonId: string): Promise<void> => {
    await updateMaster((c) => ({ ...c, regularLessons: c.regularLessons.filter((l) => l.id !== lessonId) }))
  }

  // --- Excel (operates on master data) ---
  const downloadTemplate = (): void => {
    // Include sample test data so users can see the expected format
    const sampleTeachers = [
      ['田中講師', '数,英', '数学メイン', 'tanaka@example.com'],
      ['佐藤講師', '英,数', '英語メイン', ''],
    ]
    const sampleStudents = [
      ['青木 太郎', '中3', 'aoki@example.com'],
      ['伊藤 花', '中2', ''],
      ['上田 陽介', '高1', ''],
    ]
    const sampleConstraints = [
      ['田中講師', '伊藤 花', '不可'],
    ]
    const sampleGradeConstraints = [
      ['佐藤講師', '高1', '不可'],
    ]
    const sampleRegularLessons = [
      ['田中講師', '青木 太郎', '', '数', '月', '1'],
    ]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '担当科目(カンマ区切り: ' + FIXED_SUBJECTS.join(',') + ')', 'メモ', 'メールアドレス'], ...sampleTeachers]), '講師')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '学年', 'メールアドレス'], ...sampleStudents]), '生徒')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒名', '種別(不可)'], ...sampleConstraints]), '制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '学年', '種別(不可)'], ...sampleGradeConstraints]), '学年制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒1名', '生徒2名(任意)', '科目', '曜日(月/火/水/木/金/土/日)', '時限番号'], ...sampleRegularLessons]), '通常授業')
    XLSX.writeFile(wb, 'テンプレート.xlsx')
  }

  const exportData = (): void => {
    if (!masterData) return
    const md = masterData
    const teacherRows = md.teachers.map((t) => [t.name, t.subjects.join(', '), t.memo, t.email ?? ''])
    const studentRows = md.students.map((s) => [s.name, s.grade, s.memo, s.email ?? ''])
    const constraintRows = md.constraints.map((c) => [
      md.teachers.find((t) => t.id === c.teacherId)?.name ?? c.teacherId,
      md.students.find((s) => s.id === c.studentId)?.name ?? c.studentId,
      c.type === 'incompatible' ? '不可' : '推奨',
    ])
    const gradeConstraintRows = (md.gradeConstraints ?? []).map((gc) => [
      md.teachers.find((t) => t.id === gc.teacherId)?.name ?? gc.teacherId,
      gc.grade,
      gc.type === 'incompatible' ? '不可' : '推奨',
    ])
    const dayNames = ['日', '月', '火', '水', '木', '金', '土']
    const regularLessonRows = md.regularLessons.map((l) => [
      md.teachers.find((t) => t.id === l.teacherId)?.name ?? l.teacherId,
      ...l.studentIds.map((id) => md.students.find((s) => s.id === id)?.name ?? id),
      ...(l.studentIds.length === 1 ? [''] : []),
      l.subject, dayNames[l.dayOfWeek] ?? '', l.slotNumber,
    ])
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '担当科目', 'メモ', 'メール'], ...teacherRows]), '講師')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '学年', 'メモ', 'メール'], ...studentRows]), '生徒')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒名', '種別'], ...constraintRows]), '制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '学年', '種別'], ...gradeConstraintRows]), '学年制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒1名', '生徒2名', '科目', '曜日', '時限'], ...regularLessonRows]), '通常授業')
    XLSX.writeFile(wb, '管理データ.xlsx')
  }

  const handleFileImport = async (file: File): Promise<void> => {
    if (!masterData) return
    const md = masterData
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })

    const importedTeachers: Teacher[] = []
    const importedStudents: Student[] = []
    const importedConstraints: PairConstraint[] = []
    const importedGradeConstraints: GradeConstraint[] = []
    const importedRegularLessons: RegularLesson[] = []

    const findTeacherId = (name: string): string | null => {
      const existing = md.teachers.find((t) => t.name === name)
      if (existing) return existing.id
      return importedTeachers.find((t) => t.name === name)?.id ?? null
    }
    const findStudentId = (name: string): string | null => {
      const existing = md.students.find((s) => s.name === name)
      if (existing) return existing.id
      return importedStudents.find((s) => s.name === name)?.id ?? null
    }

    const teacherWs = wb.Sheets['講師']
    if (teacherWs) {
      const rows = XLSX.utils.sheet_to_json(teacherWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const name = String(row?.[0] ?? '').trim()
        if (!name) continue
        const subjects = String(row?.[1] ?? '').split(/[、,]/).map((s) => s.trim()).filter((s) => FIXED_SUBJECTS.includes(s))
        const memo = String(row?.[2] ?? '').trim()
        const email = String(row?.[3] ?? '').trim()
        if (md.teachers.some((t) => t.name === name)) continue
        importedTeachers.push({ id: createId(), name, email, subjects, memo })
      }
    }

    const studentWs = wb.Sheets['生徒']
    if (studentWs) {
      const rows = XLSX.utils.sheet_to_json(studentWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const name = String(row?.[0] ?? '').trim()
        if (!name) continue
        const grade = String(row?.[1] ?? '').trim()
        const email = String(row?.[2] ?? '').trim()
        if (md.students.some((s) => s.name === name)) continue
        importedStudents.push({ id: createId(), name, email, grade, subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo: '', submittedAt: 0 })
      }
    }

    const constraintWs = wb.Sheets['制約']
    if (constraintWs) {
      const rows = XLSX.utils.sheet_to_json(constraintWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const tn = String(row?.[0] ?? '').trim(); const sn = String(row?.[1] ?? '').trim(); const ts = String(row?.[2] ?? '').trim()
        const tid = findTeacherId(tn); const sid = findStudentId(sn)
        if (!tid || !sid) continue
        const type: ConstraintType = ts === '推奨' ? 'recommended' : 'incompatible'
        if (md.constraints.some((c) => c.teacherId === tid && c.studentId === sid)) continue
        importedConstraints.push({ id: createId(), teacherId: tid, studentId: sid, type })
      }
    }

    const gradeConstraintWs = wb.Sheets['学年制約']
    if (gradeConstraintWs) {
      const rows = XLSX.utils.sheet_to_json(gradeConstraintWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const tn = String(row?.[0] ?? '').trim(); const grade = String(row?.[1] ?? '').trim(); const ts = String(row?.[2] ?? '').trim()
        const tid = findTeacherId(tn)
        if (!tid || !grade) continue
        const type: ConstraintType = ts === '推奨' ? 'recommended' : 'incompatible'
        if ((md.gradeConstraints ?? []).some((c) => c.teacherId === tid && c.grade === grade)) continue
        importedGradeConstraints.push({ id: createId(), teacherId: tid, grade, type })
      }
    }

    const dayNameMap: Record<string, number> = { '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 }
    const regularWs = wb.Sheets['通常授業']
    if (regularWs) {
      const rows = XLSX.utils.sheet_to_json(regularWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const tName = String(row?.[0] ?? '').trim(); const s1 = String(row?.[1] ?? '').trim(); const s2 = String(row?.[2] ?? '').trim()
        const subject = String(row?.[3] ?? '').trim(); const dayStr = String(row?.[4] ?? '').trim(); const slotNum = Number(row?.[5])
        const tid = findTeacherId(tName)
        if (!tid || !subject) continue
        const sids = [findStudentId(s1), findStudentId(s2)].filter(Boolean) as string[]
        if (sids.length === 0) continue
        const dow = dayNameMap[dayStr]
        if (dow === undefined || Number.isNaN(slotNum) || slotNum < 1) continue
        // Dedup: skip if same regular lesson already exists in master data
        const sortedSids = [...sids].sort()
        const isDup = md.regularLessons.some((existing) =>
          existing.teacherId === tid &&
          existing.subject === subject &&
          existing.dayOfWeek === dow &&
          existing.slotNumber === slotNum &&
          existing.studentIds.length === sortedSids.length &&
          [...existing.studentIds].sort().every((id, j) => id === sortedSids[j]),
        )
        if (isDup) continue
        importedRegularLessons.push({ id: createId(), teacherId: tid, studentIds: sids, subject, dayOfWeek: dow, slotNumber: slotNum })
      }
    }

    // --- Validation: check for inconsistencies ---
    const validationErrors: string[] = []
    // All teachers (existing + imported)
    const allTeachers = [...md.teachers, ...importedTeachers]
    const allStudents = [...md.students, ...importedStudents]
    // Check regular lessons
    for (const rl of importedRegularLessons) {
      const teacher = allTeachers.find((t) => t.id === rl.teacherId)
      if (teacher && !teacher.subjects.includes(rl.subject)) {
        const dayNames2 = ['日', '月', '火', '水', '木', '金', '土']
        validationErrors.push(`通常授業: ${teacher.name} の担当科目に「${rl.subject}」がありません（${dayNames2[rl.dayOfWeek]}曜${rl.slotNumber}限）`)
      }
      for (const sid of rl.studentIds) {
        const student = allStudents.find((s) => s.id === sid)
        if (!student) {
          validationErrors.push(`通常授業: 生徒ID「${sid}」が見つかりません`)
        }
      }
    }
    // Check constraints reference valid people
    for (const c of importedConstraints) {
      if (!allTeachers.some((t) => t.id === c.teacherId)) {
        validationErrors.push(`制約: 講師ID「${c.teacherId}」が見つかりません`)
      }
      if (!allStudents.some((s) => s.id === c.studentId)) {
        validationErrors.push(`制約: 生徒ID「${c.studentId}」が見つかりません`)
      }
    }
    // Check grade constraints reference valid teachers
    for (const gc of importedGradeConstraints) {
      if (!allTeachers.some((t) => t.id === gc.teacherId)) {
        validationErrors.push(`学年制約: 講師ID「${gc.teacherId}」が見つかりません`)
      }
    }

    if (validationErrors.length > 0) {
      alert(`⚠️ 取り込みエラー:\n\n${validationErrors.join('\n')}\n\nデータを修正してから再度取り込んでください。`)
      return
    }

    const added: string[] = []
    if (importedTeachers.length) added.push(`講師${importedTeachers.length}名`)
    if (importedStudents.length) added.push(`生徒${importedStudents.length}名`)
    if (importedConstraints.length) added.push(`制約${importedConstraints.length}件`)
    if (importedGradeConstraints.length) added.push(`学年制約${importedGradeConstraints.length}件`)
    if (importedRegularLessons.length) added.push(`通常授業${importedRegularLessons.length}件`)
    if (added.length === 0) { alert('新規データがありませんでした（同名は重複スキップ）。'); return }
    if (!window.confirm(`以下を取り込みます:\n${added.join(', ')}\n\nよろしいですか？`)) return

    await updateMaster((c) => ({
      ...c,
      teachers: [...c.teachers, ...importedTeachers],
      students: [...c.students, ...importedStudents],
      constraints: [...c.constraints, ...importedConstraints],
      gradeConstraints: [...(c.gradeConstraints ?? []), ...importedGradeConstraints],
      regularLessons: [...c.regularLessons, ...importedRegularLessons],
    }))
    setTimeout(() => alert('取り込み完了！'), 50)
  }

  // --- Session management ---
  const cleanupLegacyDevSession = async (): Promise<void> => {
    const legacyDev = await loadSession('dev')
    if (!legacyDev) return
    await deleteSession('dev')
  }

  const onUnlock = async (): Promise<void> => {
    saveAdminPassword(adminPassword)
    setUnlocked(true)
    await cleanupLegacyDevSession()
  }

  const onCreateSession = async (): Promise<void> => {
    const id = newSessionId.trim()
    if (!id) return
    const isMendanSession = newTerm.includes('mendan')
    if (!masterData) {
      alert('管理データが読み込まれていません。')
      return
    }
    if (isMendanSession) {
      if ((masterData.managers ?? []).length === 0 || masterData.students.length === 0) {
        alert('管理データ（マネージャー・生徒）が未登録です。先に管理データを登録してください。')
        return
      }
    } else {
      if (masterData.teachers.length === 0 && masterData.students.length === 0) {
        alert('管理データ（講師・生徒）が未登録です。先に管理データを登録してください。')
        return
      }
    }
    if (sessions.some((s) => s.id === id)) {
      alert('同じIDの特別講習が既に存在します。別のIDにしてください。')
      return
    }
    if (!newStartDate || !newEndDate) {
      alert('講習期間（開始日・終了日）を入力してください。')
      return
    }
    const seed = emptySession()
    seed.settings.name = newSessionName.trim() || id
    seed.settings.adminPassword = adminPassword
    seed.settings.startDate = newStartDate
    seed.settings.endDate = newEndDate
    seed.settings.submissionStartDate = newSubmissionStart
    seed.settings.submissionEndDate = newSubmissionEnd
    seed.settings.deskCount = newDeskCount
    seed.settings.holidays = [...newHolidays]
    seed.settings.sessionType = isMendanSession ? 'mendan' : 'lecture'
    if (isMendanSession) {
      seed.settings.slotsPerDay = 10 // 10:00-20:00 range; managers define their own times
      seed.settings.mendanStartHour = 10
    } else {
      seed.settings.slotsPerDay = newDeskCount > 0 ? 5 : 5
    }
    seed.managers = masterData.managers ?? []
    seed.teachers = masterData.teachers
    seed.students = masterData.students.map((s) => ({
      ...s, subjects: isMendanSession ? ['面談'] : [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], submittedAt: 0,
    }))
    seed.constraints = masterData.constraints
    seed.gradeConstraints = isMendanSession ? [] : masterData.gradeConstraints
    seed.regularLessons = isMendanSession ? [] : masterData.regularLessons
    try {
      const verified = await saveAndVerify(id, seed)
      if (!verified) {
        alert('特別講習の作成に失敗しました。Firebaseのセキュリティルールを確認してください。')
      }
    } catch (e) {
      alert(`特別講習の作成に失敗しました:\n${e instanceof Error ? e.message : String(e)}`)
    }
  }

  const openAdmin = (sessionId: string): void => {
    saveAdminPassword(adminPassword)
    navigate(`/admin/${sessionId}`)
  }

  const handleDeleteSession = async (sessionId: string, sessionName: string): Promise<void> => {
    const confirmed = window.confirm(`特別講習「${sessionName || sessionId}」を削除しますか？\nこの操作は元に戻せません。`)
    if (!confirmed) return
    const password = window.prompt('削除パスワードを入力してください:')
    if (password !== adminPassword) {
      alert('パスワードが正しくありません。')
      return
    }
    await deleteSession(sessionId)
    alert('特別講習を削除しました。')
  }

  const formatDate = (ms: number): string => {
    if (!ms) return '-'
    const d = new Date(ms)
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
  }

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>講習コマ割りアプリ</h2>
        <p className="muted">管理データ（講師・生徒・制約）はここで一元管理し、特別講習ごとに希望コマ数とコマ割りを管理します。</p>

        {!unlocked ? (
          <>
            <h3>管理者パスワード</h3>
            <div className="row">
              <input type="password" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} placeholder="管理者パスワード" />
              <button className="btn" type="button" onClick={() => void onUnlock()}>続行</button>
            </div>
            <p className="muted">現在は初期値を保存済みのため、入力不要で続行できます。</p>
          </>
        ) : !masterData ? (
          <div className="panel">
            <div className="loading-container">
              <div className="loading-spinner" />
              <div className="loading-text">管理データを読み込み中...</div>
            </div>
          </div>
        ) : (
          <>
            {/* --- Session management --- */}
            <div className="panel">
              <h3>新規特別講習を追加</h3>
              {(!masterData || (
                newTerm.includes('mendan')
                  ? ((masterData.managers ?? []).length === 0 && masterData.students.length === 0)
                  : (masterData.teachers.length === 0 && masterData.students.length === 0)
              )) ? (
                <p style={{ color: '#dc2626', fontWeight: 600 }}>⚠ 管理データ（{newTerm.includes('mendan') ? 'マネージャー・生徒' : '講師・生徒'}）が未登録のため、特別講習を追加できません。先に下部の管理データを登録してください。</p>
              ) : (
                <>
                  <p className="muted">作成時にマスターデータ（講師・生徒・制約・通常授業）が自動コピーされます。</p>
                  <div className="row">
                    <input value={newYear} onChange={(e) => setNewYear(e.target.value)} placeholder="西暦" style={{ width: 80 }} />
                    <select value={newTerm} onChange={(e) => setNewTerm(e.target.value as typeof newTerm)}>
                      <optgroup label="講習">
                        <option value="spring">春期講習</option>
                        <option value="summer">夏期講習</option>
                        <option value="winter">冬期講習</option>
                      </optgroup>
                      <optgroup label="面談">
                        <option value="spring-mendan">春期面談</option>
                        <option value="summer-mendan">夏期面談</option>
                        <option value="winter-mendan">冬期面談</option>
                      </optgroup>
                    </select>
                    <input value={newSessionId} onChange={(e) => setNewSessionId(e.target.value)} placeholder="ID (例: 2026-summer)" style={{ width: 160 }} />
                    <input value={newSessionName} onChange={(e) => setNewSessionName(e.target.value)} placeholder="表示名 (例: 2026 夏期講習)" />
                  </div>
                  <div className="row" style={{ marginTop: '8px', flexWrap: 'wrap', gap: '8px' }}>
                    <label className="muted" style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      講習期間:
                      <input type="date" value={newStartDate} onChange={(e) => setNewStartDate(e.target.value)} />
                      〜
                      <input type="date" value={newEndDate} onChange={(e) => setNewEndDate(e.target.value)} />
                    </label>
                  </div>
                  <div className="row" style={{ marginTop: '8px', flexWrap: 'wrap', gap: '8px' }}>
                    <label className="muted" style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      提出期間:
                      <input type="date" value={newSubmissionStart} onChange={(e) => setNewSubmissionStart(e.target.value)} />
                      〜
                      <input type="date" value={newSubmissionEnd} onChange={(e) => setNewSubmissionEnd(e.target.value)} />
                    </label>
                    <span className="muted" style={{ fontSize: '11px' }}>※この期間のみ希望URLが有効になります</span>
                  </div>
                  <div className="row" style={{ marginTop: '8px', flexWrap: 'wrap', gap: '8px' }}>
                    <label className="muted" style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      机の数:
                      <input type="number" min={0} value={newDeskCount} onChange={(e) => setNewDeskCount(Math.max(0, Number(e.target.value) || 0))} style={{ width: '60px' }} />
                      <span style={{ fontSize: '11px' }}>0=無制限</span>
                    </label>
                  </div>
                  {newTerm.includes('mendan') && (
                    <div className="row" style={{ marginTop: '8px', flexWrap: 'wrap', gap: '8px' }}>
                      <span className="muted" style={{ fontSize: '12px' }}>※ 面談時間帯はマネージャーが希望入力時に日ごとに指定します</span>
                    </div>
                  )}
                  <div className="row" style={{ marginTop: '8px', flexWrap: 'wrap', gap: '8px', alignItems: 'center' }}>
                    <span className="muted">休日:</span>
                    <input
                      type="date"
                      onChange={(e) => {
                        const val = e.target.value
                        if (val && !newHolidays.includes(val)) {
                          setNewHolidays((prev) => [...prev, val].sort())
                        }
                        e.target.value = ''
                      }}
                    />
                    {newHolidays.map((h) => (
                      <span key={h} className="badge warn" style={{ cursor: 'pointer' }} onClick={() => setNewHolidays((prev) => prev.filter((d) => d !== h))}>
                        {formatShortDate(h)} ×
                      </span>
                    ))}
                  </div>
                  <div className="row" style={{ marginTop: '12px' }}>
                    <button className="btn" type="button" onClick={() => void onCreateSession()}>特別講習を作成</button>
                  </div>
                </>
              )}
            </div>

            <div className="panel">
              <div className="row" style={{ justifyContent: 'space-between' }}>
                <h3>特別講習一覧（新しい順）</h3>
                <button className="btn secondary" type="button" onClick={() => setUnlocked(false)}>ロック</button>
              </div>
              <div className="row" style={{ marginBottom: '8px', gap: '8px' }}>
                <button className="btn secondary" type="button" onClick={() => {
                  // Backup: export all sessions as JSON
                  const payload = { sessions: sessions.map((s) => s.id), exportedAt: Date.now() }
                  const promises = sessions.map((s) => loadSession(s.id).then((d) => ({ id: s.id, data: d })))
                  void Promise.all(promises).then((results) => {
                    const backup = { ...payload, data: Object.fromEntries(results.map((r) => [r.id, r.data])) }
                    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' })
                    const url = URL.createObjectURL(blob)
                    const a = document.createElement('a')
                    a.href = url
                    a.download = `特別講習バックアップ_${new Date().toISOString().slice(0, 10)}.json`
                    a.click()
                    URL.revokeObjectURL(url)
                  })
                }}>📥 全データバックアップ</button>
                <button className="btn secondary" type="button" onClick={() => {
                  const input = document.createElement('input')
                  input.type = 'file'
                  input.accept = '.json'
                  input.onchange = async () => {
                    const file = input.files?.[0]
                    if (!file) return
                    try {
                      const text = await file.text()
                      const backup = JSON.parse(text) as { data: Record<string, SessionData> }
                      if (!backup.data || typeof backup.data !== 'object') {
                        alert('バックアップファイルの形式が正しくありません。')
                        return
                      }
                      const ids = Object.keys(backup.data)
                      const existingIds = sessions.map((s) => s.id)
                      const newIds = ids.filter((id) => !existingIds.includes(id))
                      const overwriteIds = ids.filter((id) => existingIds.includes(id))
                      let msg = `取り込み対象: ${ids.length}件の特別講習`
                      if (newIds.length > 0) msg += `\n  新規: ${newIds.join(', ')}`
                      if (overwriteIds.length > 0) msg += `\n  上書き: ${overwriteIds.join(', ')}`
                      if (!window.confirm(msg + '\n\n取り込みますか？')) return
                      for (const [id, data] of Object.entries(backup.data)) {
                        if (data) await saveSession(id, data as SessionData)
                      }
                      alert(`${ids.length}件の特別講習を取り込みました。`)
                    } catch (e) {
                      alert(`取り込みエラー: ${e instanceof Error ? e.message : String(e)}`)
                    }
                  }
                  input.click()
                }}>📤 バックアップ取り込み</button>
              </div>
              <table className="table">
                <thead><tr><th>ID</th><th>名称</th><th>作成</th><th>更新</th><th /><th /></tr></thead>
                <tbody>
                  {sessions.map((s) => (
                    <tr key={s.id}>
                      <td>{s.id}</td><td>{s.name}</td><td>{formatDate(s.createdAt)}</td><td>{formatDate(s.updatedAt)}</td>
                      <td><button className="btn" type="button" onClick={() => openAdmin(s.id)}>管理</button></td>
                      <td><button className="btn secondary" type="button" style={{ color: '#dc2626' }} onClick={() => void handleDeleteSession(s.id, s.name)}>削除</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* --- Master data management --- */}
                <div className="panel" style={{ position: 'relative' }}>
                  <h3>管理データ — Excel</h3>
                  <div className="row">
                    <button className="btn" type="button" onClick={downloadTemplate}>テンプレートExcel出力</button>
                    <button className="btn" type="button" onClick={exportData}>現状データエクセル出力</button>
                    <button className="btn secondary" type="button" onClick={() => fileInputRef.current?.click()}>Excel取り込み</button>
                    <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }}
                      onChange={(e) => { const file = e.target.files?.[0]; if (file) void handleFileImport(file); e.target.value = '' }} />
                  </div>
                </div>

                <div className="panel">
                  <h3>マネージャー登録</h3>
                  <div className="row">
                    <input value={managerName} onChange={(e) => setManagerName(e.target.value)} placeholder="マネージャー名" />
                    <input value={managerEmail} onChange={(e) => setManagerEmail(e.target.value)} placeholder="メールアドレス" type="email" />
                    <button className="btn" type="button" onClick={() => void addManager()}>追加</button>
                  </div>
                  <table className="table">
                    <thead><tr><th>名前</th><th>メール</th><th>操作</th></tr></thead>
                    <tbody>
                      {(masterData.managers ?? []).map((m) => (
                        editingManagerId === m.id ? (
                          <tr key={m.id}>
                            <td>{m.name}</td>
                            <td><input value={editManagerEmail} onChange={(e) => setEditManagerEmail(e.target.value)} type="email" /></td>
                            <td>
                              <button className="btn" type="button" onClick={() => void saveEditManager()}>保存</button>
                              <button className="btn secondary" type="button" onClick={() => setEditingManagerId(null)} style={{ marginLeft: 4 }}>キャンセル</button>
                            </td>
                          </tr>
                        ) : (
                          <tr key={m.id}><td>{m.name}</td><td>{m.email || '-'}</td>
                            <td>
                              <button className="btn" type="button" onClick={() => startEditManager(m)} style={{ marginRight: 4 }}>編集</button>
                              <button className="btn secondary" type="button" onClick={() => void removeManager(m.id)}>削除</button>
                            </td>
                          </tr>
                        )
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>講師登録</h3>
                  <div className="row">
                    <input value={teacherName} onChange={(e) => setTeacherName(e.target.value)} placeholder="講師名" />
                    <input value={teacherEmail} onChange={(e) => setTeacherEmail(e.target.value)} placeholder="メールアドレス" type="email" />
                    <select onChange={(e) => { const v = e.target.value; if (v && !teacherSubjects.includes(v)) setTeacherSubjects((p) => [...p, v]); e.target.value = '' }}>
                      <option value="">担当科目を追加</option>
                      {FIXED_SUBJECTS.filter((s) => !teacherSubjects.includes(s)).map((s) => (<option key={s} value={s}>{s}</option>))}
                    </select>
                    <input value={teacherMemo} onChange={(e) => setTeacherMemo(e.target.value)} placeholder="メモ" />
                    <button className="btn" type="button" onClick={() => void addTeacher()}>追加</button>
                  </div>
                  <div className="row">
                    {teacherSubjects.map((s) => (
                      <span key={s} className="badge ok" style={{ cursor: 'pointer' }} onClick={() => setTeacherSubjects((p) => p.filter((x) => x !== s))}>{s} ×</span>
                    ))}
                  </div>
                  <table className="table">
                    <thead><tr><th>名前</th><th>メール</th><th>科目</th><th>メモ</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.teachers.map((t) => (
                        editingTeacherId === t.id ? (
                          <tr key={t.id}>
                            <td>{t.name}</td>
                            <td><input value={editTeacherEmail} onChange={(e) => setEditTeacherEmail(e.target.value)} type="email" /></td>
                            <td>
                              <select onChange={(e) => { const v = e.target.value; if (v && !editTeacherSubjects.includes(v)) setEditTeacherSubjects((p) => [...p, v]); e.target.value = '' }}>
                                <option value="">科目追加</option>
                                {FIXED_SUBJECTS.filter((s) => !editTeacherSubjects.includes(s)).map((s) => (<option key={s} value={s}>{s}</option>))}
                              </select>
                              <div>{editTeacherSubjects.map((s) => (
                                <span key={s} className="badge ok" style={{ cursor: 'pointer' }} onClick={() => setEditTeacherSubjects((p) => p.filter((x) => x !== s))}>{s} ×</span>
                              ))}</div>
                            </td>
                            <td><input value={editTeacherMemo} onChange={(e) => setEditTeacherMemo(e.target.value)} /></td>
                            <td>
                              <button className="btn" type="button" onClick={() => void saveEditTeacher()}>保存</button>
                              <button className="btn secondary" type="button" onClick={() => setEditingTeacherId(null)} style={{ marginLeft: 4 }}>キャンセル</button>
                            </td>
                          </tr>
                        ) : (
                          <tr key={t.id}><td>{t.name}</td><td>{t.email || '-'}</td><td>{t.subjects.join(', ')}</td><td>{t.memo}</td>
                            <td>
                              <button className="btn" type="button" onClick={() => startEditTeacher(t)} style={{ marginRight: 4 }}>編集</button>
                              <button className="btn secondary" type="button" onClick={() => void removeTeacher(t.id)}>削除</button>
                            </td>
                          </tr>
                        )
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>生徒登録</h3>
                  <div className="row">
                    <input value={studentName} onChange={(e) => setStudentName(e.target.value)} placeholder="生徒名" />
                    <input value={studentEmail} onChange={(e) => setStudentEmail(e.target.value)} placeholder="メールアドレス" type="email" />
                    <select value={studentGrade} onChange={(e) => setStudentGrade(e.target.value)}>
                      <option value="">学年を選択</option>
                      {GRADE_OPTIONS.map((g) => (<option key={g} value={g}>{g}</option>))}
                    </select>
                    <button className="btn" type="button" onClick={() => void addStudent()}>追加</button>
                  </div>
                  <table className="table">
                    <thead><tr><th>名前</th><th>メール</th><th>学年</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.students.map((s) => (
                        editingStudentId === s.id ? (
                          <tr key={s.id}>
                            <td>{s.name}</td>
                            <td><input value={editStudentEmail} onChange={(e) => setEditStudentEmail(e.target.value)} type="email" /></td>
                            <td>
                              <select value={editStudentGrade} onChange={(e) => setEditStudentGrade(e.target.value)}>
                                <option value="">学年を選択</option>
                                {GRADE_OPTIONS.map((g) => (<option key={g} value={g}>{g}</option>))}
                              </select>
                            </td>
                            <td>
                              <button className="btn" type="button" onClick={() => void saveEditStudent()}>保存</button>
                              <button className="btn secondary" type="button" onClick={() => setEditingStudentId(null)} style={{ marginLeft: 4 }}>キャンセル</button>
                            </td>
                          </tr>
                        ) : (
                          <tr key={s.id}><td>{s.name}</td><td>{s.email || '-'}</td><td>{s.grade}</td>
                            <td>
                              <button className="btn" type="button" onClick={() => startEditStudent(s)} style={{ marginRight: 4 }}>編集</button>
                              <button className="btn secondary" type="button" onClick={() => void removeStudent(s.id)}>削除</button>
                            </td>
                          </tr>
                        )
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>通常授業管理</h3>
                  <div className="row">
                    <select value={regularTeacherId} onChange={(e) => setRegularTeacherId(e.target.value)}>
                      <option value="">講師を選択</option>
                      {masterData.teachers.map((t) => (<option key={t.id} value={t.id}>{t.name}</option>))}
                    </select>
                    <select value={regularStudent1Id} onChange={(e) => setRegularStudent1Id(e.target.value)}>
                      <option value="">生徒1を選択</option>
                      {masterData.students.map((s) => (<option key={s.id} value={s.id} disabled={s.id === regularStudent2Id}>{s.name}</option>))}
                    </select>
                    <select value={regularStudent2Id} onChange={(e) => setRegularStudent2Id(e.target.value)}>
                      <option value="">生徒2(任意)</option>
                      {masterData.students.map((s) => (<option key={s.id} value={s.id} disabled={s.id === regularStudent1Id}>{s.name}</option>))}
                    </select>
                    <select value={regularSubject} onChange={(e) => setRegularSubject(e.target.value)}>
                      <option value="">科目を選択</option>
                      {FIXED_SUBJECTS.map((s) => (<option key={s} value={s}>{s}</option>))}
                    </select>
                    <select value={regularDayOfWeek} onChange={(e) => setRegularDayOfWeek(e.target.value)}>
                      <option value="">曜日を選択</option>
                      <option value="0">日曜</option><option value="1">月曜</option><option value="2">火曜</option>
                      <option value="3">水曜</option><option value="4">木曜</option><option value="5">金曜</option><option value="6">土曜</option>
                    </select>
                    <input type="number" value={regularSlotNumber} onChange={(e) => setRegularSlotNumber(e.target.value)} placeholder="時限番号" min="1" />
                    <button className="btn" type="button" onClick={() => void addRegularLesson()}>追加</button>
                  </div>
                  <p className="muted">通常授業は該当する曜日・時限のスロットに最優先で割り当てられます。</p>
                  <table className="table">
                    <thead><tr><th>講師</th><th>生徒</th><th>科目</th><th>曜日</th><th>時限</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.regularLessons.map((l) => {
                        const dayNames = ['日', '月', '火', '水', '木', '金', '土']
                        return (
                          <tr key={l.id}>
                            <td>{masterData.teachers.find((t) => t.id === l.teacherId)?.name ?? '-'}</td>
                            <td>{l.studentIds.map((id) => masterData.students.find((s) => s.id === id)?.name ?? '-').join(', ')}</td>
                            <td>{l.subject}</td><td>{dayNames[l.dayOfWeek]}曜</td><td>{l.slotNumber}限</td>
                            <td><button className="btn secondary" type="button" onClick={() => void removeRegularLesson(l.id)}>削除</button></td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>講師×生徒 制約</h3>
                  <div className="row">
                    <select value={constraintTeacherId} onChange={(e) => setConstraintTeacherId(e.target.value)}>
                      <option value="">講師を選択</option>
                      {masterData.teachers.map((t) => (<option key={t.id} value={t.id}>{t.name}</option>))}
                    </select>
                    <select value={constraintStudentId} onChange={(e) => setConstraintStudentId(e.target.value)}>
                      <option value="">生徒を選択</option>
                      {masterData.students.map((s) => (<option key={s.id} value={s.id}>{s.name}</option>))}
                    </select>
                    <select value={constraintType} onChange={(e) => setConstraintType(e.target.value as ConstraintType)}>
                      <option value="incompatible">組み合わせ不可</option>
                    </select>
                    <button className="btn" type="button" onClick={() => void upsertConstraint()}>保存</button>
                  </div>
                  <table className="table">
                    <thead><tr><th>講師</th><th>生徒</th><th>種別</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.constraints.map((c) => (
                        <tr key={c.id}>
                          <td>{masterData.teachers.find((t) => t.id === c.teacherId)?.name ?? '-'}</td>
                          <td>{masterData.students.find((s) => s.id === c.studentId)?.name ?? '-'}</td>
                          <td><span className="badge warn">不可</span></td>
                          <td><button className="btn secondary" type="button" onClick={() => void removeConstraint(c.id)}>削除</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  <h4 style={{ marginTop: '16px' }}>講師×学年 制約</h4>
                  <div className="row" style={{ flexWrap: 'wrap', gap: '8px', alignItems: 'center' }}>
                    <select value={gradeConstraintTeacherId} onChange={(e) => setGradeConstraintTeacherId(e.target.value)}>
                      <option value="">講師を選択</option>
                      {masterData.teachers.map((t) => (<option key={t.id} value={t.id}>{t.name}</option>))}
                    </select>
                    <select value={gradeConstraintGrade} onChange={(e) => setGradeConstraintGrade(e.target.value)}>
                      <option value="">学年を選択</option>
                      {GRADE_OPTIONS.map((g) => (<option key={g} value={g}>{g}</option>))}
                    </select>
                    <select value={gradeConstraintType} onChange={(e) => setGradeConstraintType(e.target.value as ConstraintType)}>
                      <option value="incompatible">担当不可</option>
                    </select>
                    <div style={{ display: 'flex', gap: '4px', flexWrap: 'wrap', alignItems: 'center' }}>
                      <span style={{ fontSize: '13px', color: '#64748b' }}>科目:</span>
                      {FIXED_SUBJECTS.map((subj) => (
                        <label key={subj} style={{ display: 'flex', alignItems: 'center', gap: '2px', fontSize: '13px', cursor: 'pointer' }}>
                          <input
                            type="checkbox"
                            checked={gradeConstraintSubjects.includes(subj)}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setGradeConstraintSubjects((prev) => [...prev, subj])
                              } else {
                                setGradeConstraintSubjects((prev) => prev.filter((s) => s !== subj))
                              }
                            }}
                          />
                          {subj}
                        </label>
                      ))}
                      <span style={{ fontSize: '11px', color: '#94a3b8' }}>未選択=全科目</span>
                    </div>
                    <button className="btn" type="button" onClick={() => void upsertGradeConstraint()}>保存</button>
                  </div>
                  <table className="table">
                    <thead><tr><th>講師</th><th>学年</th><th>科目</th><th>種別</th><th>操作</th></tr></thead>
                    <tbody>
                      {(masterData.gradeConstraints ?? []).map((gc) => (
                        <tr key={gc.id}>
                          <td>{masterData.teachers.find((t) => t.id === gc.teacherId)?.name ?? '-'}</td>
                          <td>{gc.grade}</td>
                          <td>{gc.subjects && gc.subjects.length > 0 ? gc.subjects.join(', ') : '全科目'}</td>
                          <td><span className="badge warn">不可</span></td>
                          <td><button className="btn secondary" type="button" onClick={() => void removeGradeConstraint(gc.id)}>削除</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
          </>
        )}
      </div>
    </div>
  )
}

// --- Data Analytics Panel ---
type AnalyticsTab = 'teacher' | 'student' | 'day' | 'subject'

const AnalyticsPanel = ({ data, slotKeys }: { data: SessionData; slotKeys: string[] }) => {
  const [tab, setTab] = useState<AnalyticsTab>('teacher')
  const dayNames = ['日', '月', '火', '水', '木', '金', '土']

  // Precompute helpers
  const allSlotAssignments = useMemo(() => {
    const entries: { slot: string; assignment: Assignment }[] = []
    for (const slot of slotKeys) {
      for (const a of data.assignments[slot] ?? []) {
        entries.push({ slot, assignment: a })
      }
    }
    return entries
  }, [data.assignments, slotKeys])

  const dateSet = useMemo(() => {
    const s = new Set<string>()
    for (const sk of slotKeys) s.add(sk.split('_')[0])
    return s
  }, [slotKeys])

  // --- Teacher Analytics ---
  const teacherStats = useMemo(() => {
    return data.teachers.map((teacher) => {
      const myAssignments = allSlotAssignments.filter((e) => e.assignment.teacherId === teacher.id)
      const totalSlots = myAssignments.length
      const regularSlots = myAssignments.filter((e) => e.assignment.isRegular).length
      const specialSlots = totalSlots - regularSlots
      const dates = new Set(myAssignments.map((e) => e.slot.split('_')[0]))
      const attendanceDays = dates.size

      // Available slots
      const availableSlots = slotKeys.filter((sk) =>
        hasAvailability(data.availability, 'teacher', teacher.id, sk),
      ).length

      // Subject breakdown (count per-student subjects)
      const subjectMap: Record<string, number> = {}
      for (const e of myAssignments) {
        if (!e.assignment.isRegular) {
          for (const sid of e.assignment.studentIds) {
            const subj = getStudentSubject(e.assignment, sid)
            subjectMap[subj] = (subjectMap[subj] ?? 0) + 1
          }
        }
      }

      // Student count (unique special students)
      const studentIds = new Set<string>()
      for (const e of myAssignments) {
        if (!e.assignment.isRegular) {
          for (const sid of e.assignment.studentIds) studentIds.add(sid)
        }
      }

      return {
        teacher,
        totalSlots,
        regularSlots,
        specialSlots,
        attendanceDays,
        availableSlots,
        utilization: availableSlots > 0 ? Math.round((totalSlots / availableSlots) * 100) : 0,
        subjectMap,
        uniqueStudentCount: studentIds.size,
      }
    })
  }, [data, allSlotAssignments, slotKeys])

  // --- Student Analytics ---
  const studentStats = useMemo(() => {
    return data.students.filter((s) => s.submittedAt > 0).map((student) => {
      const myAssignments = allSlotAssignments.filter((e) =>
        e.assignment.studentIds.includes(student.id),
      )
      const totalSlots = myAssignments.filter((e) => !e.assignment.isRegular).length
      const regularSlots = myAssignments.filter((e) => e.assignment.isRegular).length
      const dates = new Set(myAssignments.map((e) => e.slot.split('_')[0]))

      // Per-subject desired vs assigned (using per-student subjects)
      const subjectDetails = Object.entries(student.subjectSlots).map(([subj, desired]) => {
        const assigned = myAssignments.filter((e) => !e.assignment.isRegular && getStudentSubject(e.assignment, student.id) === subj).length
        return { subject: subj, desired, assigned, diff: assigned - desired }
      })

      const totalDesired = Object.values(student.subjectSlots).reduce((s, c) => s + c, 0)
      const totalAssigned = totalSlots

      // Unique teachers
      const teacherIds = new Set<string>()
      for (const e of myAssignments) {
        if (!e.assignment.isRegular) teacherIds.add(e.assignment.teacherId)
      }

      return {
        student,
        totalSlots,
        regularSlots,
        totalDesired,
        totalAssigned,
        fulfillment: totalDesired > 0 ? Math.round((totalAssigned / totalDesired) * 100) : 0,
        attendanceDays: dates.size,
        subjectDetails,
        uniqueTeacherCount: teacherIds.size,
      }
    })
  }, [data, allSlotAssignments])

  // --- Day of Week Analytics ---
  const dayStats = useMemo(() => {
    const daysUsed = new Map<number, { dates: Set<string>; totalPairs: number; regularPairs: number; specialPairs: number; teacherIds: Set<string>; studentIds: Set<string> }>()
    for (let d = 0; d < 7; d++) daysUsed.set(d, { dates: new Set(), totalPairs: 0, regularPairs: 0, specialPairs: 0, teacherIds: new Set(), studentIds: new Set() })
    for (const e of allSlotAssignments) {
      const date = e.slot.split('_')[0]
      const dow = getIsoDayOfWeek(date)
      const entry = daysUsed.get(dow)!
      entry.dates.add(date)
      entry.totalPairs++
      if (e.assignment.isRegular) entry.regularPairs++
      else entry.specialPairs++
      if (e.assignment.teacherId) entry.teacherIds.add(e.assignment.teacherId)
      for (const sid of e.assignment.studentIds) entry.studentIds.add(sid)
    }
    return [1, 2, 3, 4, 5, 6, 0].map((d) => ({ dayOfWeek: d, dayName: dayNames[d], ...daysUsed.get(d)! }))
      .filter((d) => d.dates.size > 0)
  }, [allSlotAssignments])

  // --- Subject Analytics ---
  const subjectStats = useMemo(() => {
    const subjectMap = new Map<string, { totalPairs: number; studentIds: Set<string>; teacherIds: Set<string>; totalDesired: number; totalAssigned: number }>()
    for (const subj of FIXED_SUBJECTS) {
      subjectMap.set(subj, { totalPairs: 0, studentIds: new Set(), teacherIds: new Set(), totalDesired: 0, totalAssigned: 0 })
    }
    for (const e of allSlotAssignments) {
      if (e.assignment.isRegular) continue
      // Count per-student subjects
      for (const sid of e.assignment.studentIds) {
        const subj = getStudentSubject(e.assignment, sid)
        const entry = subjectMap.get(subj)
        if (!entry) continue
        entry.totalAssigned++
        entry.studentIds.add(sid)
        if (e.assignment.teacherId) entry.teacherIds.add(e.assignment.teacherId)
      }
      // Count unique pairs per primary subject
      const primaryEntry = subjectMap.get(e.assignment.subject)
      if (primaryEntry) primaryEntry.totalPairs++
    }
    for (const student of data.students) {
      for (const [subj, desired] of Object.entries(student.subjectSlots)) {
        const entry = subjectMap.get(subj)
        if (entry) entry.totalDesired += desired
      }
    }
    return FIXED_SUBJECTS.map((subj) => {
      const entry = subjectMap.get(subj)!
      return { subject: subj, ...entry }
    }).filter((s) => s.totalPairs > 0 || s.totalDesired > 0)
  }, [data, allSlotAssignments])

  // --- Date-level detail ---
  const dateStats = useMemo(() => {
    const result: { date: string; dayName: string; totalPairs: number; regularPairs: number; specialPairs: number; maxSlotPairs: number }[] = []
    for (const date of [...dateSet].sort()) {
      const dow = getIsoDayOfWeek(date)
      let totalPairs = 0; let regularPairs = 0; let specialPairs = 0; let maxSlotPairs = 0
      for (let s = 1; s <= data.settings.slotsPerDay; s++) {
        const pairs = data.assignments[`${date}_${s}`] ?? []
        const cnt = pairs.length
        totalPairs += cnt
        regularPairs += pairs.filter((a) => a.isRegular).length
        specialPairs += pairs.filter((a) => !a.isRegular).length
        if (cnt > maxSlotPairs) maxSlotPairs = cnt
      }
      result.push({ date, dayName: dayNames[dow], totalPairs, regularPairs, specialPairs, maxSlotPairs })
    }
    return result
  }, [data, dateSet])

  const tabs: { key: AnalyticsTab; label: string }[] = [
    { key: 'teacher', label: '講師別' },
    { key: 'student', label: '生徒別' },
    { key: 'day', label: '曜日・日別' },
    { key: 'subject', label: '科目別' },
  ]

  return (
    <div className="panel analytics-panel">
      <div className="row" style={{ marginBottom: '12px', gap: '4px' }}>
        {tabs.map((t) => (
          <button
            key={t.key}
            className={`btn${tab === t.key ? '' : ' secondary'}`}
            type="button"
            onClick={() => setTab(t.key)}
            style={{ fontSize: '0.85em', padding: '5px 12px' }}
          >
            {t.label}
          </button>
        ))}
      </div>

      {tab === 'teacher' && (
        <div style={{ overflowX: 'auto' }}>
          <table className="table">
            <thead>
              <tr>
                <th>講師名</th>
                <th>出勤日数</th>
                <th>出席可能コマ</th>
                <th>通常授業</th>
                <th>特別講習</th>
                <th>合計コマ</th>
                <th>稼働率</th>
                <th>担当科目内訳</th>
                <th>担当生徒数</th>
              </tr>
            </thead>
            <tbody>
              {teacherStats.map((ts) => (
                <tr key={ts.teacher.id}>
                  <td style={{ fontWeight: 500 }}>{ts.teacher.name}</td>
                  <td>{ts.attendanceDays}日</td>
                  <td>{ts.availableSlots}</td>
                  <td>{ts.regularSlots}</td>
                  <td>{ts.specialSlots}</td>
                  <td style={{ fontWeight: 600 }}>{ts.totalSlots}</td>
                  <td>
                    <span style={{
                      color: ts.utilization >= 80 ? '#16a34a' : ts.utilization >= 50 ? '#d97706' : '#dc2626',
                      fontWeight: 600,
                    }}>
                      {ts.utilization}%
                    </span>
                  </td>
                  <td style={{ fontSize: '0.85em' }}>
                    {Object.entries(ts.subjectMap).map(([subj, cnt]) => `${subj}:${cnt}`).join(' ') || '-'}
                  </td>
                  <td>{ts.uniqueStudentCount}名</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {tab === 'student' && (
        <div style={{ overflowX: 'auto' }}>
          <table className="table">
            <thead>
              <tr>
                <th>生徒名</th>
                <th>学年</th>
                <th>希望計</th>
                <th>割当計</th>
                <th>充足率</th>
                <th>科目別（希望/割当）</th>
                <th>通常</th>
                <th>出席日数</th>
                <th>担当講師数</th>
              </tr>
            </thead>
            <tbody>
              {studentStats.map((ss) => (
                <tr key={ss.student.id}>
                  <td style={{ fontWeight: 500 }}>{ss.student.name}</td>
                  <td>{ss.student.grade}</td>
                  <td>{ss.totalDesired}</td>
                  <td style={{ fontWeight: 600 }}>{ss.totalAssigned}</td>
                  <td>
                    <span style={{
                      color: ss.fulfillment >= 100 ? '#16a34a' : ss.fulfillment >= 70 ? '#d97706' : '#dc2626',
                      fontWeight: 600,
                    }}>
                      {ss.fulfillment}%
                    </span>
                  </td>
                  <td style={{ fontSize: '0.85em' }}>
                    {ss.subjectDetails.map((sd) => {
                      const color = sd.diff > 0 ? '#dc2626' : sd.diff < 0 ? '#d97706' : '#16a34a'
                      return (
                        <span key={sd.subject} style={{ marginRight: '8px' }}>
                          {sd.subject}:<span style={{ color, fontWeight: 500 }}>{sd.desired}/{sd.assigned}</span>
                        </span>
                      )
                    })}
                    {ss.subjectDetails.length === 0 && '-'}
                  </td>
                  <td>{ss.regularSlots}</td>
                  <td>{ss.attendanceDays}日</td>
                  <td>{ss.uniqueTeacherCount}名</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {tab === 'day' && (
        <>
          <h4 style={{ margin: '0 0 8px' }}>曜日別サマリー</h4>
          <div style={{ overflowX: 'auto', marginBottom: '16px' }}>
            <table className="table">
              <thead>
                <tr>
                  <th>曜日</th>
                  <th>日数</th>
                  <th>通常</th>
                  <th>特別</th>
                  <th>合計ペア</th>
                  <th>講師数</th>
                  <th>生徒数</th>
                </tr>
              </thead>
              <tbody>
                {dayStats.map((ds) => (
                  <tr key={ds.dayOfWeek}>
                    <td style={{ fontWeight: 500 }}>{ds.dayName}曜</td>
                    <td>{ds.dates.size}日</td>
                    <td>{ds.regularPairs}</td>
                    <td>{ds.specialPairs}</td>
                    <td style={{ fontWeight: 600 }}>{ds.totalPairs}</td>
                    <td>{ds.teacherIds.size}名</td>
                    <td>{ds.studentIds.size}名</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <h4 style={{ margin: '0 0 8px' }}>日別詳細</h4>
          <div style={{ overflowX: 'auto' }}>
            <table className="table">
              <thead>
                <tr>
                  <th>日付</th>
                  <th>曜日</th>
                  <th>通常</th>
                  <th>特別</th>
                  <th>合計ペア</th>
                  <th>最大同時ペア</th>
                </tr>
              </thead>
              <tbody>
                {dateStats.map((ds) => (
                  <tr key={ds.date}>
                    <td style={{ fontWeight: 500 }}>{formatShortDate(ds.date)}</td>
                    <td>{ds.dayName}</td>
                    <td>{ds.regularPairs}</td>
                    <td>{ds.specialPairs}</td>
                    <td style={{ fontWeight: 600 }}>{ds.totalPairs}</td>
                    <td>
                      {ds.maxSlotPairs}
                      {(data.settings.deskCount ?? 0) > 0 && (
                        <span style={{ color: ds.maxSlotPairs >= (data.settings.deskCount ?? 0) ? '#dc2626' : '#64748b', fontSize: '0.85em' }}>
                          /{data.settings.deskCount}
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === 'subject' && (
        <div style={{ overflowX: 'auto' }}>
          <table className="table">
            <thead>
              <tr>
                <th>科目</th>
                <th>全生徒希望計</th>
                <th>割当済ペア</th>
                <th>充足率</th>
                <th>担当講師数</th>
                <th>受講生徒数</th>
              </tr>
            </thead>
            <tbody>
              {subjectStats.map((ss) => {
                const fulfillment = ss.totalDesired > 0 ? Math.round((ss.totalAssigned / ss.totalDesired) * 100) : 0
                return (
                  <tr key={ss.subject}>
                    <td style={{ fontWeight: 600, fontSize: '1.1em' }}>{ss.subject}</td>
                    <td>{ss.totalDesired}コマ</td>
                    <td>{ss.totalPairs}ペア（{ss.totalAssigned}人回）</td>
                    <td>
                      <span style={{
                        color: fulfillment >= 100 ? '#16a34a' : fulfillment >= 70 ? '#d97706' : '#dc2626',
                        fontWeight: 600,
                      }}>
                        {fulfillment}%
                      </span>
                    </td>
                    <td>{ss.teacherIds.size}名</td>
                    <td>{ss.studentIds.size}名</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}

const AdminPage = () => {
  const { sessionId = 'main' } = useParams()
  const navigate = useNavigate()
  const location = useLocation()
  const skipAuth = (location.state as { skipAuth?: boolean } | null)?.skipAuth === true
  const { data, setData, loading, error: sessionError } = useSessionData(sessionId)
  const [authorized, setAuthorized] = useState(import.meta.env.DEV || skipAuth)
  const [recentlyUpdated, setRecentlyUpdated] = useState<Set<string>>(new Set())
  const [dragInfo, setDragInfo] = useState<{ sourceSlot: string; sourceIdx: number; teacherId: string; studentIds: string[]; studentDragId?: string; studentDragSubject?: string } | null>(null)
  const [showAnalytics, setShowAnalytics] = useState(false)
  const [showRules, setShowRules] = useState(false)
  const [emailSendLog, setEmailSendLog] = useState<Record<string, { time: string; type: string }>>({})

  const prevSnapshotRef = useRef<{ availability: Record<string, string[]>; studentSubmittedAt: Record<string, number> } | null>(null)
  const masterSyncDoneRef = useRef(false)

  // Sync master data (regularLessons, constraints) into session on first load
  useEffect(() => {
    if (!data || masterSyncDoneRef.current) return
    masterSyncDoneRef.current = true
    const isMendanSession = data.settings.sessionType === 'mendan'
    loadMasterData().then((master) => {
      if (!master) return
      const needsUpdate =
        JSON.stringify(data.constraints) !== JSON.stringify(master.constraints) ||
        JSON.stringify(data.gradeConstraints) !== JSON.stringify(isMendanSession ? [] : master.gradeConstraints) ||
        JSON.stringify(data.regularLessons) !== JSON.stringify(isMendanSession ? [] : master.regularLessons)
      if (!needsUpdate) return
      const next: SessionData = {
        ...data,
        constraints: master.constraints,
        gradeConstraints: isMendanSession ? [] : (master.gradeConstraints ?? []),
        regularLessons: isMendanSession ? [] : master.regularLessons,
      }
      setData(next)
      saveSession(sessionId, next).catch(() => { /* ignore */ })
    }).catch(() => { /* ignore */ })
  }, [data, sessionId])

  // Track real-time changes to show "just updated" indicators for teachers/students
  useEffect(() => {
    if (!data) return
    const prev = prevSnapshotRef.current
    if (!prev) {
      // First load — just record snapshot, don't show indicators
      prevSnapshotRef.current = {
        availability: { ...data.availability },
        studentSubmittedAt: Object.fromEntries(data.students.map((s) => [s.id, s.submittedAt])),
      }
      return
    }
    const updatedIds: string[] = []
    // Check teachers/managers: availability change
    for (const teacher of data.teachers) {
      const key = `teacher:${teacher.id}`
      const prevSlots = (prev.availability[key] ?? []).slice().sort().join(',')
      const currSlots = (data.availability[key] ?? []).slice().sort().join(',')
      if (prevSlots !== currSlots) updatedIds.push(teacher.id)
    }
    for (const manager of (data.managers ?? [])) {
      const key = `manager:${manager.id}`
      const prevSlots = (prev.availability[key] ?? []).slice().sort().join(',')
      const currSlots = (data.availability[key] ?? []).slice().sort().join(',')
      if (prevSlots !== currSlots) updatedIds.push(manager.id)
    }
    // Check students: submittedAt change
    for (const student of data.students) {
      const prevAt = prev.studentSubmittedAt[student.id] ?? 0
      if (student.submittedAt !== prevAt) updatedIds.push(student.id)
    }
    // Update snapshot
    prevSnapshotRef.current = {
      availability: { ...data.availability },
      studentSubmittedAt: Object.fromEntries(data.students.map((s) => [s.id, s.submittedAt])),
    }
    if (updatedIds.length > 0) {
      setRecentlyUpdated((p) => new Set([...p, ...updatedIds]))
      // Clear after 3 seconds
      const timer = setTimeout(() => {
        setRecentlyUpdated((p) => {
          const next = new Set(p)
          for (const id of updatedIds) next.delete(id)
          return next
        })
      }, 3000)
      return () => clearTimeout(timer)
    }
  }, [data])
  const buildInputUrl = (personType: PersonType, personId: string): string => {
    const base = window.location.origin + (import.meta.env.BASE_URL ?? '/').replace(/\/$/, '')
    return `${base}/#/availability/${sessionId}/${personType}/${personId}`
  }

  const copyInputUrl = async (personType: PersonType, personId: string): Promise<void> => {
    const url = buildInputUrl(personType, personId)
    try {
      await navigator.clipboard.writeText(url)
      alert('URLをコピーしました')
    } catch {
      window.prompt('URLをコピーしてください:', url)
    }
  }

  const EMAIL_TYPE_LABELS: Record<string, string> = {
    'input-request': '予定入力依頼',
    'confirmed': 'コマ割り確定',
    'changed': 'コマ割り変更',
  }

  const buildMailtoForPerson = (person: { id: string; name: string; email: string }, personType: PersonType, contentType: string): string => {
    const sessionName = data?.settings.name ?? ''
    const url = buildInputUrl(personType, person.id)
    let subject: string
    let bodyLines: string[]
    switch (contentType) {
      case 'confirmed':
        subject = `【${sessionName}】コマ割り確定のお知らせ`
        bodyLines = [
          `${person.name} 様`,
          '',
          `${sessionName}のコマ割りが確定しましたのでお知らせします。`,
          '',
          '以下のURLから確定したスケジュールをご確認ください。',
          '',
          url,
        ]
        break
      case 'changed':
        subject = `【${sessionName}】コマ割り変更のお知らせ`
        bodyLines = [
          `${person.name} 様`,
          '',
          `${sessionName}のコマ割りに変更がありましたのでお知らせします。`,
          '',
          '以下のURLから最新のスケジュールをご確認ください。',
          '',
          url,
        ]
        break
      default: // input-request
        subject = `【${sessionName}】希望入力URLのご案内`
        bodyLines = [
          `${person.name} 様`,
          '',
          `${sessionName}の希望入力URLをお送りします。`,
          '',
          '以下のURLからご自身の希望を入力してください。',
          '',
          url,
        ]
        break
    }
    return `mailto:${encodeURIComponent(person.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyLines.join('\n'))}`
  }

  const handleEmailSend = (person: { id: string; name: string; email: string }, personType: PersonType): void => {
    const contentType = data?.settings.confirmed ? 'confirmed' : 'input-request'
    const typeLabel = EMAIL_TYPE_LABELS[contentType] ?? '予定入力依頼'
    const now = new Date()
    const timeStr = now.toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    // Log the send
    setEmailSendLog((prev) => ({ ...prev, [person.id]: { time: timeStr, type: typeLabel } }))
    // Download receipt PDF first, then open mailto after a short delay
    const sentAt = now.toLocaleString('ja-JP')
    const mailtoUrl = buildMailtoForPerson(person, personType, contentType)
    void downloadEmailReceiptPdf({
      sessionName: data?.settings.name ?? '',
      recipientName: person.name,
      emailType: typeLabel,
      sentAt,
    }).finally(() => {
      // Open mailto after PDF is generated (or on error)
      setTimeout(() => { window.location.href = mailtoUrl }, 300)
    })
  }

  const openInputPage = async (personType: PersonType, personId: string): Promise<void> => {
    navigate(`/availability/${sessionId}/${personType}/${personId}`, { state: { fromAdminInput: true } })
  }

  useEffect(() => {
    setAuthorized(import.meta.env.DEV || skipAuth)
  }, [sessionId, skipAuth])

  const slotKeys = useMemo(() => (data ? buildSlotKeys(data.settings) : []), [data])

  // Mendan (interview) mode: managers act as instructors instead of teachers
  const isMendan = data?.settings.sessionType === 'mendan'
  const mendanStart = data?.settings.mendanStartHour ?? 10

  // For mendan: compute which slot numbers actually have manager availability
  const mendanActiveSlots = useMemo(() => {
    if (!isMendan || !data) return new Set<number>()
    const active = new Set<number>()
    for (const manager of (data.managers ?? [])) {
      const key = personKey('manager', manager.id)
      for (const sk of (data.availability[key] ?? [])) {
        const slotNum = Number(sk.split('_')[1])
        if (!isNaN(slotNum)) active.add(slotNum)
      }
    }
    return active
  }, [isMendan, data])

  // For mendan: filtered slotKeys that only include manager-available slots
  const effectiveSlotKeys = useMemo(() => {
    if (!isMendan || mendanActiveSlots.size === 0) return slotKeys
    return slotKeys.filter((sk) => {
      const slotNum = Number(sk.split('_')[1])
      return mendanActiveSlots.has(slotNum)
    })
  }, [isMendan, slotKeys, mendanActiveSlots])

  const instructors: Teacher[] = useMemo(() => {
    if (!data) return []
    if (isMendan) {
      return (data.managers ?? []).map((m) => ({
        id: m.id, name: m.name, email: m.email, subjects: ['面談'], memo: '',
      }))
    }
    return data.teachers
  }, [data, isMendan])
  const instructorPersonType: PersonType = isMendan ? 'manager' : 'teacher'
  const instructorLabel = isMendan ? 'マネージャー' : '講師'

  const persist = async (next: SessionData): Promise<void> => {
    setData(next)
    await saveSession(sessionId, next)
  }

  const update = async (updater: (current: SessionData) => SessionData): Promise<void> => {
    if (!data) return
    await persist(updater(data))
  }

  // --- Undo / Redo for assignments ---
  const undoStackRef = useRef<Record<string, Assignment[]>[]>([])
  const redoStackRef = useRef<Record<string, Assignment[]>[]>([])
  const [undoCount, setUndoCount] = useState(0)
  const [redoCount, setRedoCount] = useState(0)
  const MAX_UNDO = 50

  const pushUndo = (assignments: Record<string, Assignment[]>): void => {
    undoStackRef.current = [...undoStackRef.current.slice(-(MAX_UNDO - 1)), JSON.parse(JSON.stringify(assignments))]
    redoStackRef.current = []
    setUndoCount(undoStackRef.current.length)
    setRedoCount(0)
  }

  const updateAssignments = async (updater: (current: SessionData) => SessionData): Promise<void> => {
    if (!data) return
    pushUndo(data.assignments)
    await persist(updater(data))
  }

  const handleUndo = async (): Promise<void> => {
    if (undoStackRef.current.length === 0 || !data) return
    const prev = undoStackRef.current.pop()!
    redoStackRef.current.push(JSON.parse(JSON.stringify(data.assignments)))
    setUndoCount(undoStackRef.current.length)
    setRedoCount(redoStackRef.current.length)
    await persist({ ...data, assignments: prev })
  }

  const handleRedo = async (): Promise<void> => {
    if (redoStackRef.current.length === 0 || !data) return
    const next = redoStackRef.current.pop()!
    undoStackRef.current.push(JSON.parse(JSON.stringify(data.assignments)))
    setUndoCount(undoStackRef.current.length)
    setRedoCount(redoStackRef.current.length)
    await persist({ ...data, assignments: next })
  }

  const createSession = async (): Promise<void> => {
    const seed = emptySession()
    try {
      const verified = await saveAndVerify(sessionId, seed)
      if (!verified) {
        alert('特別講習の作成に失敗しました。Firebaseのセキュリティルールを確認してください。')
      }
    } catch (e) {
      alert(`特別講習の作成に失敗しました:\n${e instanceof Error ? e.message : String(e)}\n\nFirebase ConsoleでAuthentication（匿名）とFirestoreルールを確認してください。`)
    }
  }

  useEffect(() => {
    if (!data) return
    if (import.meta.env.DEV || skipAuth) { setAuthorized(true); return }
    const password = readSavedAdminPassword()
    setAuthorized(password === data.settings.adminPassword)
  }, [data, skipAuth])

  // Auto-sync master data on session open (preserves settings, availability, assignments)
  const masterSyncedRef = useRef('')
  useEffect(() => {
    if (!data || !authorized || masterSyncedRef.current === sessionId) return
    masterSyncedRef.current = sessionId
    void (async () => {
      const master = await loadMasterData()
      if (!master) return
      const mergedStudents = master.students.map((ms) => {
        const existing = data.students.find((s) => s.id === ms.id)
        if (existing) {
          return { ...ms, subjects: existing.subjects, subjectSlots: existing.subjectSlots, unavailableDates: existing.unavailableDates, preferredSlots: existing.preferredSlots ?? [], unavailableSlots: existing.unavailableSlots ?? [], submittedAt: existing.submittedAt }
        }
        return { ...ms, subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], submittedAt: 0 }
      })
      // Preserve settings, availability, and assignments — only update people/constraints/regularLessons
      const next: SessionData = {
        ...data,
        managers: master.managers ?? [],
        teachers: master.teachers,
        students: mergedStudents,
        constraints: master.constraints,
        gradeConstraints: master.gradeConstraints,
        regularLessons: master.regularLessons,
      }
      // Only save if something actually changed
      const changed =
        JSON.stringify(data.managers ?? []) !== JSON.stringify(next.managers) ||
        JSON.stringify(data.teachers) !== JSON.stringify(next.teachers) ||
        JSON.stringify(data.students.map((s) => ({ id: s.id, name: s.name, grade: s.grade, memo: s.memo }))) !==
          JSON.stringify(next.students.map((s) => ({ id: s.id, name: s.name, grade: s.grade, memo: s.memo }))) ||
        JSON.stringify(data.constraints) !== JSON.stringify(next.constraints) ||
        JSON.stringify(data.gradeConstraints) !== JSON.stringify(next.gradeConstraints) ||
        JSON.stringify(data.regularLessons) !== JSON.stringify(next.regularLessons)
      if (changed) {
        setData(next)
        await saveSession(sessionId, next)
      }
    })()
  }, [data, authorized]) // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-fill regular lessons when date range or regular lessons change (skip for mendan)
  const regularFillSigRef = useRef('')
  useEffect(() => {
    if (!data || !authorized || slotKeys.length === 0) return
    if (data.settings.sessionType === 'mendan') return
    const assignmentStateSig = slotKeys
      .map((slot) => (data.assignments[slot] ?? []).map((a) => assignmentSignature(a)).sort().join(';'))
      .join(',')
    const sig = `${slotKeys.join(',')}|${data.regularLessons.map((l) => `${l.id}:${l.dayOfWeek}:${l.slotNumber}:${l.teacherId}:${l.studentIds.join('+')}:${l.subject}`).join(',')}|${assignmentStateSig}`
    if (sig === regularFillSigRef.current) return
    regularFillSigRef.current = sig

    let changed = false
    const nextAssignments = { ...data.assignments }

    for (const slot of slotKeys) {
      const existing = nextAssignments[slot]
      const slotRegularLessons = findRegularLessonsForSlot(data.regularLessons, slot)

      if (slotRegularLessons.length === 0) {
        if (existing && existing.length > 0 && existing.every((a) => a.isRegular)) {
          delete nextAssignments[slot]
          changed = true
        }
        continue
      }

      // Don't overwrite manual (non-regular) assignments
      if (existing && existing.some((a) => hasMeaningfulManualAssignment(a))) continue

      const expectedRegulars = slotRegularLessons.map((lesson) => ({
        teacherId: lesson.teacherId,
        studentIds: lesson.studentIds,
        subject: lesson.subject,
        isRegular: true,
      }))

      const expectedSig = expectedRegulars
        .map((a) => assignmentSignature(a))
        .sort()
      const existingSig = (existing ?? [])
        .filter((a) => a.isRegular)
        .map((a) => assignmentSignature(a))
        .sort()

      if (expectedSig.length === existingSig.length && expectedSig.every((sigItem, idx) => sigItem === existingSig[idx])) {
        continue
      }

      nextAssignments[slot] = expectedRegulars
      changed = true
    }

    // Auto-register regular lesson slots as teacher availability
    const nextAvailability = { ...data.availability }
    let availChanged = false
    for (const lesson of data.regularLessons) {
      const teacherKey = personKey('teacher', lesson.teacherId)
      const currentSlots = new Set(nextAvailability[teacherKey] ?? [])
      for (const sk of slotKeys) {
        const dayOfWeek = getSlotDayOfWeek(sk)
        const slotNumber = getSlotNumber(sk)
        if (lesson.dayOfWeek === dayOfWeek && lesson.slotNumber === slotNumber) {
          if (!currentSlots.has(sk)) {
            currentSlots.add(sk)
            availChanged = true
          }
        }
      }
      if (availChanged || currentSlots.size !== (data.availability[teacherKey] ?? []).length) {
        nextAvailability[teacherKey] = [...currentSlots]
      }
    }

    if (changed || availChanged) {
      void persist({ ...data, assignments: nextAssignments, availability: nextAvailability })
    }
  }, [slotKeys, data, authorized]) // eslint-disable-line react-hooks/exhaustive-deps

  const applyAutoAssign = async (): Promise<void> => {
    if (!data) return

    // Mendan FCFS auto-assign
    if (isMendan) {
      const { assignments: nextAssignments, unassignedParents } = buildMendanAutoAssignments(data, slotKeys)
      const submittedCount = data.students.filter((s) => s.submittedAt > 0).length
      const assignedCount = submittedCount - unassignedParents.length

      await updateAssignments((current) => ({
        ...current,
        assignments: nextAssignments,
        autoAssignHighlights: { added: {}, changed: {} },
        settings: { ...current.settings, lastAutoAssignedAt: Date.now() },
      }))

      const msg = unassignedParents.length > 0
        ? `自動割当完了: ${assignedCount}/${submittedCount}名を割当しました。\n\n未割当の保護者:\n${unassignedParents.join('\n')}`
        : `自動割当完了: 提出済み${assignedCount}名全員を割当しました。`
      alert(msg)
      return
    }

    // Lecture auto-assign (existing logic)
    const { assignments: nextAssignments, changeLog, changedPairSignatures, addedPairSignatures, changeDetails } = buildIncrementalAutoAssignments(data, slotKeys)

    const highlightAdded: Record<string, string[]> = {}
    const highlightChanged: Record<string, string[]> = {}
    const highlightDetails: Record<string, Record<string, string>> = {}
    const allSlotSet = new Set([...Object.keys(addedPairSignatures), ...Object.keys(changedPairSignatures)])
    for (const slot of allSlotSet) {
      if ((addedPairSignatures[slot] ?? []).length > 0) highlightAdded[slot] = [...addedPairSignatures[slot]]
      if ((changedPairSignatures[slot] ?? []).length > 0) highlightChanged[slot] = [...changedPairSignatures[slot]]
      if (changeDetails[slot]) highlightDetails[slot] = { ...changeDetails[slot] }
    }

    const remainingStudents = data.students
      .map((student) => {
        const remaining = Object.entries(student.subjectSlots)
          .map(([subject, desired]) => {
            const assigned = countStudentSubjectLoad(nextAssignments, student.id, subject)
            return { subject, remaining: desired - assigned }
          })
          .filter((item) => item.remaining > 0)
        if (remaining.length === 0) return null
        return { studentName: student.name, remaining }
      })
      .filter(Boolean) as { studentName: string; remaining: { subject: string; remaining: number }[] }[]

    const overAssignedStudents = data.students
      .map((student) => {
        const over = Object.entries(student.subjectSlots)
          .map(([subject, desired]) => {
            const assigned = countStudentSubjectLoad(nextAssignments, student.id, subject)
            return { subject, over: assigned - desired }
          })
          .filter((item) => item.over > 0)
        if (over.length === 0) return null
        return { studentName: student.name, over }
      })
      .filter(Boolean) as { studentName: string; over: { subject: string; over: number }[] }[]

    const shortageEntries = collectTeacherShortages(data, nextAssignments)

    await updateAssignments((current) => ({
      ...current,
      assignments: nextAssignments,
      autoAssignHighlights: { added: highlightAdded, changed: highlightChanged, changeDetails: highlightDetails },
      settings: { ...current.settings, lastAutoAssignedAt: Date.now() },
    }))
    const remainingMessage = remainingStudents.length > 0
      ? `\n\n未充足の生徒:\n${remainingStudents
          .map((item) => `${item.studentName}: ${item.remaining.map((r) => `${r.subject}残${r.remaining}`).join(', ')}`)
          .join('\n')}`
      : ''
    const overAssignedMessage = overAssignedStudents.length > 0
      ? `\n\n過割当の生徒:\n${overAssignedStudents
          .map((item) => `${item.studentName}: ${item.over.map((r) => `${r.subject}+${r.over}`).join(', ')}`)
          .join('\n')}`
      : ''
    const shortageMessage = shortageEntries.length > 0
      ? `\n\n講師不足:\n${shortageEntries
          .map((item) => `${slotLabel(item.slot, isMendan, mendanStart)} ${item.detail}`)
          .join('\n')}`
      : ''
    if (changeLog.length > 0) {
      alert(`自動提案完了: ${changeLog.length}件の変更がありました。${remainingMessage}${overAssignedMessage}${shortageMessage}`)
    } else {
      alert(`自動提案完了: 変更はありませんでした。${remainingMessage}${overAssignedMessage}${shortageMessage}`)
    }
  }

  const resetAssignments = async (): Promise<void> => {
    if (!window.confirm('コマ割りをリセットしますか？\n（手動割当と自動提案結果を全てクリアします）')) return
    await updateAssignments((current) => ({
      ...current,
      assignments: {},
      autoAssignHighlights: { added: {}, changed: {} },
    }))
  }

  const setSlotTeacher = async (slot: string, idx: number, teacherId: string): Promise<void> => {
    await updateAssignments((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      if (!teacherId) {
        slotAssignments.splice(idx, 1)
        const nextAssignments = { ...current.assignments }
        if (slotAssignments.length === 0) {
          delete nextAssignments[slot]
        } else {
          nextAssignments[slot] = slotAssignments
        }
        return { ...current, assignments: nextAssignments }
      }

      const prev = slotAssignments[idx]
      const currentInstructor = isMendan
        ? (current.managers ?? []).find((m) => m.id === teacherId)
        : current.teachers.find((item) => item.id === teacherId)
      const instructorSubjects = isMendan ? ['面談'] : ((currentInstructor && 'subjects' in currentInstructor) ? (currentInstructor as Teacher).subjects : [])
      const nextSubject =
        prev?.subject && instructorSubjects.includes(prev.subject)
          ? prev.subject
          : (instructorSubjects[0] ?? '')

      slotAssignments[idx] = {
        teacherId,
        studentIds: prev?.teacherId === teacherId ? prev.studentIds : [],
        subject: nextSubject,
      }

      return {
        ...current,
        assignments: { ...current.assignments, [slot]: slotAssignments },
      }
    })
  }

  const addSlotAssignment = async (slot: string): Promise<void> => {
    await updateAssignments((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      const deskCount = current.settings.deskCount ?? 0
      if (deskCount > 0 && slotAssignments.length >= deskCount) {
        alert(`机の数(${deskCount})の上限に達しています。`)
        return current
      }
      slotAssignments.push({ teacherId: '', studentIds: [], subject: '' })
      return {
        ...current,
        assignments: { ...current.assignments, [slot]: slotAssignments },
      }
    })
  }

  const setSlotStudent = async (slot: string, idx: number, position: number, studentId: string): Promise<void> => {
    await updateAssignments((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      const assignment = slotAssignments[idx]
      if (!assignment) {
        return current
      }

      const prevIds = [...assignment.studentIds]
      const removedId = prevIds[position]
      if (studentId === '') {
        prevIds.splice(position, 1)
      } else {
        prevIds[position] = studentId
      }
      const studentIds = prevIds.filter(Boolean)

      const teacher = current.teachers.find((item) => item.id === assignment.teacherId)

      // Build per-student subjects map
      const prevStudentSubjects = assignment.studentSubjects ?? {}
      const newStudentSubjects: Record<string, string> = {}
      for (const sid of studentIds) {
        if (sid === studentId && !prevStudentSubjects[sid]) {
          // New student: pick a viable subject from teacher's subjects
          const student = current.students.find((s) => s.id === sid)
          const viable = (teacher?.subjects ?? []).filter((subj) => student?.subjects.includes(subj))
          newStudentSubjects[sid] = viable[0] ?? assignment.subject
        } else {
          // Keep existing subject, validate it's still available
          const existingSubj = prevStudentSubjects[sid] ?? assignment.subject
          const student = current.students.find((s) => s.id === sid)
          const teacherCanTeach = teacher?.subjects.includes(existingSubj) ?? false
          const studentCanLearn = student?.subjects.includes(existingSubj) ?? false
          if (teacherCanTeach && studentCanLearn) {
            newStudentSubjects[sid] = existingSubj
          } else {
            const viable = (teacher?.subjects ?? []).filter((subj) => student?.subjects.includes(subj))
            newStudentSubjects[sid] = viable[0] ?? existingSubj
          }
        }
      }
      // Clean up removed student
      if (removedId && !studentIds.includes(removedId)) {
        delete newStudentSubjects[removedId]
      }

      // Determine the primary subject (first student's subject)
      const primarySubject = studentIds.length > 0
        ? (newStudentSubjects[studentIds[0]] ?? assignment.subject)
        : assignment.subject

      slotAssignments[idx] = {
        ...assignment,
        studentIds,
        subject: primarySubject,
        studentSubjects: studentIds.length > 0 ? newStudentSubjects : undefined,
      }

      return {
        ...current,
        assignments: { ...current.assignments, [slot]: slotAssignments },
      }
    })
  }

  /** Set subject for a specific student within an assignment pair */
  const setSlotSubject = async (slot: string, idx: number, subject: string, studentId?: string): Promise<void> => {
    await updateAssignments((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      const assignment = slotAssignments[idx]
      if (!assignment) {
        return current
      }
      if (studentId) {
        // Per-student subject change
        const studentSubjects = { ...(assignment.studentSubjects ?? {}) }
        studentSubjects[studentId] = subject
        // Update primary subject to first student's subject
        const primarySubject = assignment.studentIds.length > 0
          ? (studentSubjects[assignment.studentIds[0]] ?? assignment.subject)
          : subject
        slotAssignments[idx] = { ...assignment, subject: primarySubject, studentSubjects }
      } else {
        // Set all students to the same subject
        const studentSubjects: Record<string, string> = {}
        for (const sid of assignment.studentIds) {
          studentSubjects[sid] = subject
        }
        slotAssignments[idx] = { ...assignment, subject, studentSubjects: assignment.studentIds.length > 0 ? studentSubjects : undefined }
      }
      return {
        ...current,
        assignments: { ...current.assignments, [slot]: slotAssignments },
      }
    })
  }

  /** Move an assignment from one slot to another (drag-and-drop) */
  const moveAssignment = async (sourceSlot: string, sourceIdx: number, targetSlot: string): Promise<void> => {
    await updateAssignments((current) => {
      const srcAssignments = [...(current.assignments[sourceSlot] ?? [])]
      const moved = srcAssignments[sourceIdx]
      if (!moved || moved.isRegular) return current

      // Desk count check
      const deskCount = current.settings.deskCount ?? 0
      const targetAssignments = [...(current.assignments[targetSlot] ?? [])]
      if (deskCount > 0 && targetAssignments.length >= deskCount) return current

      // Check teacher not already in target
      if (moved.teacherId && targetAssignments.some((a) => a.teacherId === moved.teacherId)) return current

      // Check students not already assigned in target slot
      if (moved.studentIds.some((sid) => targetAssignments.some((a) => a.studentIds.includes(sid)))) return current

      // Check teacher has availability for target slot
      if (moved.teacherId && !hasAvailability(current.availability, instructorPersonType, moved.teacherId, targetSlot)) return current

      // Check all assigned students/parents are available in target slot
      const movedStudents = current.students.filter((s) => moved.studentIds.includes(s.id))
      if (movedStudents.some((student) =>
        isMendan ? !isParentAvailableForMendan(current.availability, student.id, targetSlot) : !isStudentAvailable(student, targetSlot),
      )) return current

      // Move
      srcAssignments.splice(sourceIdx, 1)
      targetAssignments.push(moved)
      const nextAssignments = { ...current.assignments }
      if (srcAssignments.length === 0) {
        delete nextAssignments[sourceSlot]
      } else {
        nextAssignments[sourceSlot] = srcAssignments
      }
      nextAssignments[targetSlot] = targetAssignments
      return { ...current, assignments: nextAssignments }
    })
  }

  /** Move a single student from one assignment to another (or to a new assignment in target slot) */
  const moveStudentToSlot = async (
    sourceSlot: string,
    sourceIdx: number,
    studentId: string,
    targetSlot: string,
    targetIdx?: number, // if provided, add to existing assignment; otherwise create new
  ): Promise<void> => {
    await updateAssignments((current) => {
      const srcAssignments = [...(current.assignments[sourceSlot] ?? [])]
      const srcAssignment = srcAssignments[sourceIdx]
      if (!srcAssignment || srcAssignment.isRegular) return current
      if (!srcAssignment.studentIds.includes(studentId)) return current

      // Check student/parent availability in target slot
      const student = current.students.find((s) => s.id === studentId)
      if (student && (isMendan
        ? !isParentAvailableForMendan(current.availability, student.id, targetSlot)
        : !isStudentAvailable(student, targetSlot))) return current

      // Check student not already in target slot
      const targetAssignments = [...(current.assignments[targetSlot] ?? [])]
      if (targetAssignments.some((a) => a.studentIds.includes(studentId))) return current

      // Remove student from source assignment
      const updatedSrcStudentIds = srcAssignment.studentIds.filter((sid) => sid !== studentId)
      const updatedSrcStudentSubjects = { ...srcAssignment.studentSubjects }
      delete updatedSrcStudentSubjects[studentId]

      const studentSubject = getStudentSubject(srcAssignment, studentId)

      if (targetIdx !== undefined && targetIdx >= 0 && targetIdx < targetAssignments.length) {
        // Add to existing target assignment (max 2 students)
        const targetAssignment = targetAssignments[targetIdx]
        if (targetAssignment.studentIds.length >= 2) return current
        if (targetAssignment.isRegular) return current

        const updatedTargetStudentIds = [...targetAssignment.studentIds, studentId]
        const updatedTargetStudentSubjects = { ...(targetAssignment.studentSubjects ?? {}) }
        // Set the student's subject in the target assignment
        updatedTargetStudentSubjects[studentId] = studentSubject
        targetAssignments[targetIdx] = {
          ...targetAssignment,
          studentIds: updatedTargetStudentIds,
          studentSubjects: updatedTargetStudentSubjects,
        }
      } else {
        // Create new assignment in target slot with just this student
        const deskCount = current.settings.deskCount ?? 0
        if (deskCount > 0 && targetAssignments.length >= deskCount) return current
        targetAssignments.push({
          teacherId: '',
          studentIds: [studentId],
          subject: studentSubject,
          studentSubjects: { [studentId]: studentSubject },
        })
      }

      // Update source: if no students left, remove the assignment entirely
      if (updatedSrcStudentIds.length === 0) {
        srcAssignments.splice(sourceIdx, 1)
      } else {
        srcAssignments[sourceIdx] = {
          ...srcAssignment,
          studentIds: updatedSrcStudentIds,
          studentSubjects: updatedSrcStudentSubjects,
        }
      }

      const nextAssignments = { ...current.assignments }
      if (srcAssignments.length === 0) {
        delete nextAssignments[sourceSlot]
      } else {
        nextAssignments[sourceSlot] = srcAssignments
      }
      nextAssignments[targetSlot] = targetAssignments
      return { ...current, assignments: nextAssignments }
    })
  }

  if (loading) {
    return (
      <div className="app-shell">
        <div className="panel">読み込み中...</div>
      </div>
    )
  }

  if (sessionError) {
    return (
      <div className="app-shell">
        <div className="panel">
          <h2>Firebaseエラー</h2>
          <p style={{ color: '#dc2626' }}>{sessionError}</p>
          <p className="muted">Firestoreのセキュリティルールを確認してください。<br />Firebase Console → Firestore Database → ルール</p>
          <pre style={{ fontSize: '11px', background: '#f3f4f6', padding: '8px', borderRadius: '4px', whiteSpace: 'pre-wrap' }}>
{`rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if request.auth != null;
    }
  }
}`}
          </pre>
          <p className="muted">上記ルールを設定し、Firebase Console → Authentication → Sign-in method で「匿名」を有効にしてください。</p>
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  if (!data) {
    return (
      <div className="app-shell">
        <div className="panel">
          <h2>特別講習: {sessionId}</h2>
          <div className="row">
            <button className="btn" type="button" onClick={createSession}>
              空の特別講習を作成
            </button>
          </div>
          <p className="muted">作成後に管理パスワードや期間を変更してください。</p>
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <div className="panel">
        <div className="row">
          <h2>管理画面: {data.settings.name} ({sessionId})</h2>
          <Link to="/" state={{ directHome: true }}>ホーム</Link>
        </div>
        <p className="muted">管理者のみ編集できます。希望入力は個別URLで配布してください。</p>
      </div>

      {!authorized ? (
        <div className="panel">
          <h3>管理者パスワードが一致しません</h3>
          <p className="muted">
            トップ画面で管理者パスワードを入力し「続行」してから、もう一度この特別講習を開いてください。
          </p>
          <Link to="/">トップへ戻る</Link>
        </div>
      ) : (
        <>
          <div className="panel">
            <h3>{instructorLabel}一覧</h3>
            <table className="table">
              <thead><tr><th>名前</th>{!isMendan && <th>科目</th>}<th>提出データ</th><th>代行入力</th><th>共有</th></tr></thead>
              <tbody>
                {instructors.map((instructor) => {
                  const submittedAt = (data.teacherSubmittedAt ?? {})[instructor.id] ?? 0
                  return (
                  <tr key={instructor.id}>
                    <td>
                      {instructor.name}
                      {recentlyUpdated.has(instructor.id) && (
                        <span className="badge ok" style={{ marginLeft: '8px', fontSize: '11px', animation: 'fadeIn 0.3s' }}>✓ 更新済</span>
                      )}
                    </td>
                    {!isMendan && <td>{instructor.subjects.join(', ')}</td>}
                    <td>
                      {submittedAt ? (
                        <span style={{ fontSize: '0.85em', color: '#16a34a' }}>
                          {new Date(submittedAt).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })}
                          {' '}提出済
                        </span>
                      ) : (
                        <span style={{ fontSize: '0.85em', color: '#dc2626', fontWeight: 600 }}>未提出</span>
                      )}
                    </td>
                    <td><button className="btn secondary" type="button" onClick={() => void openInputPage(instructorPersonType, instructor.id)}>入力ページ</button></td>
                    <td>
                      {instructor.email
                        ? <>
                            <button className="btn secondary" type="button" onClick={() => handleEmailSend(instructor, instructorPersonType)}>✉ メール送信</button>
                            {emailSendLog[instructor.id] && (
                              <span style={{ fontSize: '0.75em', color: '#2563eb', marginLeft: 6 }}>
                                {emailSendLog[instructor.id].time} {emailSendLog[instructor.id].type} 送信済み
                              </span>
                            )}
                          </>
                        : <button className="btn secondary" type="button" onClick={() => void copyInputUrl(instructorPersonType, instructor.id)}>URLコピー</button>
                      }
                    </td>
                  </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <h3>{isMendan ? '保護者一覧' : '生徒一覧'}</h3>
            <p className="muted">{isMendan ? '面談可能時間帯は保護者が希望URLから入力します。' : '希望コマ数・不可日は生徒本人が希望URLから入力します。'}</p>
            <table className="table">
              <thead><tr><th>名前</th>{!isMendan && <th>学年</th>}<th>提出データ</th><th>代行入力</th><th>共有</th></tr></thead>
              <tbody>
                {data.students.map((student) => (
                  <tr key={student.id}>
                    <td>
                      {student.name}{isMendan ? ' 保護者' : ''}
                      {recentlyUpdated.has(student.id) && (
                        <span className="badge ok" style={{ marginLeft: '8px', fontSize: '11px', animation: 'fadeIn 0.3s' }}>✓ 更新済</span>
                      )}
                    </td>
                    {!isMendan && <td>{student.grade}</td>}
                    <td>
                      {student.submittedAt ? (
                        <span style={{ fontSize: '0.85em', color: '#16a34a' }}>
                          {new Date(student.submittedAt).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })}
                          {' '}提出済
                        </span>
                      ) : (
                        <span style={{ fontSize: '0.85em', color: '#dc2626', fontWeight: 600 }}>未提出</span>
                      )}
                    </td>
                    <td><button className="btn secondary" type="button" onClick={() => void openInputPage('student', student.id)}>入力ページ</button></td>
                    <td>
                      {student.email
                        ? <>
                            <button className="btn secondary" type="button" onClick={() => handleEmailSend(student, 'student')}>✉ メール送信</button>
                            {emailSendLog[student.id] && (
                              <span style={{ fontSize: '0.75em', color: '#2563eb', marginLeft: 6 }}>
                                {emailSendLog[student.id].time} {emailSendLog[student.id].type} 送信済み
                              </span>
                            )}
                          </>
                        : <button className="btn secondary" type="button" onClick={() => void copyInputUrl('student', student.id)}>URLコピー</button>
                      }
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <div className="row">
              <h3>コマ割り</h3>
              {(() => {
                if (isMendan) {
                  // Mendan: show simple assigned/unassigned parent counts
                  const assignedParentIds = new Set<string>()
                  for (const slotAssignments of Object.values(data.assignments)) {
                    for (const a of slotAssignments) {
                      for (const sid of a.studentIds) assignedParentIds.add(sid)
                    }
                  }
                  const submittedParents = data.students.filter((s) => s.submittedAt > 0)
                  const assignedCount = submittedParents.filter((s) => assignedParentIds.has(s.id)).length
                  const unassignedCount = submittedParents.length - assignedCount
                  const totalParents = data.students.length
                  const unsubmittedCount = totalParents - submittedParents.length

                  return (
                    <>
                      {assignedCount === 0 ? (
                        <span className="badge" style={{ background: '#e5e7eb', color: '#374151' }}>未割当</span>
                      ) : unassignedCount > 0 ? (
                        <span className="badge warn" title={`提出済み${submittedParents.length}名中、${unassignedCount}名が未割当`} style={{ cursor: 'help' }}>
                          未割当: {unassignedCount}名
                        </span>
                      ) : (
                        <span className="badge ok">提出済み{assignedCount}名全員割当完了</span>
                      )}
                      {unsubmittedCount > 0 && (
                        <span className="badge" style={{ background: '#fff7ed', color: '#c2410c', border: '1px solid #fed7aa', cursor: 'help' }} title={`${unsubmittedCount}名の保護者が未提出`}>
                          未提出: {unsubmittedCount}名
                        </span>
                      )}
                    </>
                  )
                }

                const studentsWithRemaining = data.students
                  .map((student) => {
                    const remaining = Object.entries(student.subjectSlots)
                      .map(([subj, desired]) => {
                        const assigned = countStudentSubjectLoad(data.assignments, student.id, subj)
                        return { subj, rem: desired - assigned }
                      })
                      .filter((r) => r.rem !== 0)
                    if (remaining.length === 0) return null
                    return { name: student.name, remaining }
                  })
                  .filter(Boolean) as { name: string; remaining: { subj: string; rem: number }[] }[]

                const underAssigned = studentsWithRemaining.filter((s) => s.remaining.some((r) => r.rem > 0))
                const overAssigned = studentsWithRemaining.filter((s) => s.remaining.some((r) => r.rem < 0))
                const teacherShortages = collectTeacherShortages(data, data.assignments)

                const underTooltip = underAssigned
                  .map((s) => `${s.name}: ${s.remaining.filter((r) => r.rem > 0).map((r) => `${r.subj}残${r.rem}`).join(', ')}`)
                  .join('\n')
                const overTooltip = overAssigned
                  .map((s) => `${s.name}: ${s.remaining.filter((r) => r.rem < 0).map((r) => `${r.subj}${r.rem}`).join(', ')}`)
                  .join('\n')
                const shortageTooltip = teacherShortages
                  .map((item) => `${slotLabel(item.slot, isMendan, mendanStart)}: ${item.detail}`)
                  .join('\n')

                const hasAnyAssignment = Object.keys(data.assignments).length > 0
                const hasAnyDesired = data.students.some((s) => Object.values(s.subjectSlots).some((v) => v > 0))

                return (
                  <>
                    {!hasAnyAssignment || !hasAnyDesired ? (
                      <span className="badge" style={{ background: '#e5e7eb', color: '#374151' }}>未割当</span>
                    ) : underAssigned.length > 0 ? (
                      <span className="badge warn" title={underTooltip} style={{ cursor: 'help' }}>
                        残コマあり: {underAssigned.length}名
                      </span>
                    ) : (
                      <span className="badge ok">全員割当完了</span>
                    )}
                    {overAssigned.length > 0 && (
                      <span className="badge" title={overTooltip} style={{ cursor: 'help', background: '#fef2f2', color: '#dc2626', border: '1px solid #fca5a5' }}>
                        過割当: {overAssigned.length}名
                      </span>
                    )}
                    {teacherShortages.length > 0 && (
                      <span className="badge" title={shortageTooltip} style={{ cursor: 'help', background: '#fff1f2', color: '#be123c', border: '1px solid #fda4af' }}>
                        {instructorLabel}不足: {teacherShortages.length}件
                      </span>
                    )}
                  </>
                )
              })()}
              <button className="btn secondary" type="button" onClick={() => void applyAutoAssign()}>
                {isMendan ? '自動割当（先着順）' : '自動提案'}
              </button>
              <button className="btn secondary" type="button" onClick={() => void handleUndo()} disabled={undoCount === 0} title="元に戻す (Undo)">
                ↩ 戻す
              </button>
              <button className="btn secondary" type="button" onClick={() => void handleRedo()} disabled={redoCount === 0} title="やり直し (Redo)">
                ↪ やり直し
              </button>
              <button className="btn secondary" type="button" onClick={() => void resetAssignments()}>
                コマ割りリセット
              </button>
              <button className="btn" type="button" onClick={() => {
                if (!data) return
                void exportSchedulePdf({
                  sessionName: data.settings.name,
                  startDate: data.settings.startDate,
                  endDate: data.settings.endDate,
                  slotsPerDay: data.settings.slotsPerDay,
                  holidays: data.settings.holidays,
                  assignments: data.assignments,
                  getTeacherName: (id) => instructors.find((t) => t.id === id)?.name ?? '',
                  getStudentName: (id) => data.students.find((s) => s.id === id)?.name ?? '',
                  getStudentSubject,
                  getIsoDayOfWeek,
                })
              }}>
                PDF出力
              </button>
              <button className={`btn${showAnalytics ? '' : ' secondary'}`} type="button" onClick={() => setShowAnalytics((v) => !v)}>
                📊 データ分析
              </button>
              <button className={`btn${showRules ? '' : ' secondary'}`} type="button" onClick={() => setShowRules((v) => !v)}>
                📖 ルール説明
              </button>
              <button
                className={`btn${data.settings.confirmed ? '' : ' secondary'}`}
                type="button"
                style={data.settings.confirmed ? { background: '#16a34a', borderColor: '#16a34a' } : {}}
                onClick={() => {
                  const next = !data.settings.confirmed
                  if (next && !confirm('コマ割りを確定しますか？\n確定すると、入力URLがカレンダー表示に切り替わります。')) return
                  void update((c) => ({ ...c, settings: { ...c.settings, confirmed: next } }))
                }}
              >
                {data.settings.confirmed ? '✅ 確定済み' : '確定する'}
              </button>
            </div>
            <p className="muted">{isMendan ? 'マネージャー1人 + 保護者1人の面談を先着順で自動割当。' : '通常授業は日付確定時に自動配置。特別講習は自動提案で割当。講師1人 + 生徒1〜2人。'}</p>
            <p className="muted" style={{ fontSize: '12px' }}>{isMendan ? 'ペアはドラッグで別コマへ移動可' : '★=通常授業　⚠=制約不可　ペアはドラッグで別コマへ移動可'}</p>
            {showRules && (
              <div className="rules-panel" style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '8px', padding: '16px 20px', marginBottom: '12px', fontSize: '14px', lineHeight: '1.8' }}>
                <h3 style={{ margin: '0 0 12px', fontSize: '16px' }}>📖 コマ割りルール</h3>
                <div style={{ display: 'grid', gap: '12px' }}>
                  {isMendan ? (
                    <>
                      <section>
                        <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>🏫 基本構成</h4>
                        <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                          <li>1コマ = <b>マネージャー1人</b> ＋ <b>保護者1人</b></li>
                          <li>同じコマに複数の面談を配置可能（机数上限あり）</li>
                          <li>各保護者は1回の面談を割当されます</li>
                        </ul>
                      </section>
                      <section>
                        <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>🤖 自動割当（先着順）</h4>
                        <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                          <li>保護者の提出順（先着）で優先的に割当します</li>
                          <li>マネージャーの空き時間帯と保護者の希望時間帯が一致するコマに割当</li>
                          <li>自動割当後、手動で調整可能です</li>
                        </ul>
                      </section>
                      <section>
                        <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>🔄 操作方法</h4>
                        <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                          <li>ペアをドラッグ＆ドロップで別のコマへ移動可能</li>
                          <li>「＋」ボタンでコマ内にペアを追加</li>
                          <li>「×」ボタンでペアを削除</li>
                          <li>PDF出力でスケジュール表を出力（A3用紙対応）</li>
                        </ul>
                      </section>
                    </>
                  ) : (
                    <>
                  <section>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>🏫 基本構成</h4>
                    <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                      <li>1コマ = <b>講師1人</b> ＋ <b>生徒1〜2人</b></li>
                      <li>同じコマに複数のペアを配置可能（机数上限あり）</li>
                      <li>同じ生徒が同じコマに重複して入ることはできません</li>
                    </ul>
                  </section>
                  <section>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>📅 通常授業（★マーク）</h4>
                    <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                      <li>マスタデータで登録した曜日・コマ番号に毎週自動配置されます</li>
                      <li>日付が確定すると自動的にスケジュールに反映</li>
                      <li>通常授業のペアは編集・移動できません（変更はマスタデータから）</li>
                    </ul>
                  </section>
                  <section>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>🤖 自動提案</h4>
                    <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                      <li>生徒の希望コマ数を元に、空きコマへ自動的に割り当てます</li>
                      <li>講師・生徒の出勤可能日、制約ルール、科目の共通性を考慮</li>
                      <li>同じ科目の生徒同士を優先的にペアにします</li>
                      <li>講師の出勤日数が少なくなるよう連続コマ配置を優先</li>
                      <li>自動提案後、手動で調整可能です</li>
                    </ul>
                  </section>
                  <section>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>⚠️ 制約ルール</h4>
                    <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                      <li><b>講師×生徒 制約</b>：特定の講師と生徒の組み合わせを不可に設定</li>
                      <li><b>講師×学年 制約</b>：特定の講師が特定学年を担当不可に設定（科目指定も可能）</li>
                      <li>制約に違反する割当は ⚠ マークで警告表示されます</li>
                      <li>手動で制約違反の割当を強制することも可能です（確認あり）</li>
                    </ul>
                  </section>
                  <section>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>📝 科目について</h4>
                    <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                      <li>講師と生徒それぞれに担当/受講科目を設定</li>
                      <li>共通の科目がある講師・生徒のみがペアになれます</li>
                      <li>2人ペアで異なる科目の組み合わせも可能です</li>
                    </ul>
                  </section>
                  <section>
                    <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>🔄 操作方法</h4>
                    <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569' }}>
                      <li>ペアをドラッグ＆ドロップで別のコマへ移動可能</li>
                      <li>「＋」ボタンでコマ内にペアを追加</li>
                      <li>「×」ボタンでペアを削除</li>
                      <li>PDF出力でスケジュール表を出力（A3用紙対応）</li>
                    </ul>
                  </section>
                    </>
                  )}
                </div>
              </div>
            )}
            {showAnalytics && <AnalyticsPanel data={data} slotKeys={isMendan ? effectiveSlotKeys : slotKeys} />}
            <div className="grid-slots">
              {(isMendan ? effectiveSlotKeys : slotKeys).map((slot) => {
                const slotAssignments = data.assignments[slot] ?? []
                const usedTeacherIds = new Set(slotAssignments.map((a) => a.teacherId).filter(Boolean))

                // D&D: compute validity of this slot as a drop target
                const deskCount = data.settings.deskCount ?? 0
                const isDragActive = dragInfo !== null
                const isStudentDrag = isDragActive && !!dragInfo.studentDragId
                const isSameSlot = isDragActive && dragInfo.sourceSlot === slot
                const isDeskFull = isDragActive && !isStudentDrag && deskCount > 0 && slotAssignments.length >= deskCount
                const isTeacherConflict = isDragActive && !isStudentDrag && dragInfo.teacherId && usedTeacherIds.has(dragInfo.teacherId)
                const draggedStudents = isDragActive ? data.students.filter((s) => dragInfo.studentIds.includes(s.id)) : []
                const hasUnavailableStudent = isDragActive && draggedStudents.some((student) =>
                  isMendan ? !isParentAvailableForMendan(data.availability, student.id, slot) : !isStudentAvailable(student, slot))
                const hasStudentConflict = isDragActive && dragInfo.studentIds.some((sid) => {
                  // For student drag within same slot, exclude the source assignment from conflict check
                  if (isStudentDrag && isSameSlot) {
                    return slotAssignments.some((a, aIdx) => aIdx !== dragInfo.sourceIdx && a.studentIds.includes(sid))
                  }
                  return slotAssignments.some((a) => a.studentIds.includes(sid))
                })
                const hasTeacherUnavailable = isDragActive && !isStudentDrag && dragInfo.teacherId ? !hasAvailability(data.availability, instructorPersonType, dragInfo.teacherId, slot) : false
                // For student drag to same slot: valid if dropping onto a different assignment within the same slot
                const isStudentDragSameSlotOk = isStudentDrag && isSameSlot
                const isDropValid = isDragActive && (!isSameSlot || isStudentDragSameSlotOk) && !isDeskFull && !isTeacherConflict && !hasUnavailableStudent && !hasStudentConflict && !hasTeacherUnavailable
                const slotDragClass = isDragActive ? (isSameSlot && !isStudentDragSameSlotOk ? '' : isDropValid ? ' drag-valid' : ' drag-invalid') : ''

                return (
                  <div className={`slot-card${slotDragClass}`} key={slot}
                    onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = isDropValid ? 'move' : 'none' }}
                    onDrop={(e) => {
                      e.preventDefault()
                      if (!isDropValid) {
                        setDragInfo(null)
                        return
                      }
                      try {
                        const raw = e.dataTransfer.getData('text/plain')
                        if (!raw) {
                          setDragInfo(null)
                          return
                        }
                        const payload = JSON.parse(raw) as { sourceSlot: string; sourceIdx: number; studentDragId?: string }
                        if (payload.studentDragId) {
                          // Student-level drag: move student to this slot as a new assignment
                          void moveStudentToSlot(payload.sourceSlot, payload.sourceIdx, payload.studentDragId, slot)
                        } else {
                          if (payload.sourceSlot === slot) {
                            setDragInfo(null)
                            return
                          }
                          void moveAssignment(payload.sourceSlot, payload.sourceIdx, slot)
                        }
                      } catch { /* ignore */ }
                      setDragInfo(null)
                    }}
                  >
                    <div className="slot-title">
                      <div>
                        {slotLabel(slot, isMendan, mendanStart)}
                        {(data.settings.deskCount ?? 0) > 0 && (
                          <span style={{ fontSize: '0.75em', color: slotAssignments.length >= (data.settings.deskCount ?? 0) ? '#dc2626' : '#6b7280', marginLeft: '6px' }}>
                            {slotAssignments.length}/{data.settings.deskCount}
                          </span>
                        )}
                      </div>
                      <button
                        className="btn secondary slot-add-btn"
                        type="button"
                        title="ペア追加"
                        onClick={() => void addSlotAssignment(slot)}
                      >
                        ＋
                      </button>
                    </div>
                    <div className="list">
                      {slotAssignments.map((assignment, idx) => {
                        const selectedTeacher = instructors.find((t) => t.id === assignment.teacherId)

                        const isIncompatiblePair = !isMendan && assignment.teacherId && data.students.filter((s) => assignment.studentIds.includes(s.id)).some((s) => {
                          const pt = constraintFor(data.constraints, assignment.teacherId, s.id)
                          const subj = getStudentSubject(assignment, s.id)
                          const gt = gradeConstraintFor(data.gradeConstraints ?? [], assignment.teacherId, s.grade, subj)
                          return pt === 'incompatible' || gt === 'incompatible'
                        })
                        const sig = assignmentSignature(assignment)
                        const hl = data.autoAssignHighlights ?? {}
                        const isAutoAdded = (hl.added?.[slot] ?? []).includes(sig)
                        const isAutoChanged = (hl.changed?.[slot] ?? []).includes(sig)
                        const isAutoDiff = isAutoAdded || isAutoChanged
                        const changeDetail = hl.changeDetails?.[slot]?.[sig] ?? ''

                        return (
                          <div
                            key={idx}
                            className={`assignment-block${assignment.isRegular ? ' assignment-block-regular' : ''}${isIncompatiblePair ? ' assignment-block-incompatible' : ''}${isAutoDiff ? ' assignment-block-auto-updated' : ''}${isDragActive && isStudentDrag && !assignment.isRegular && assignment.studentIds.length < 2 ? ' assignment-block-drop-target' : ''}`}
                            draggable={!assignment.isRegular}
                            onDragStart={(e) => {
                              // Student row drag is handled separately with stopPropagation
                              if ((e.target as HTMLElement).closest('.student-draggable')) {
                                return // already handled by student row
                              }
                              const payload = JSON.stringify({ sourceSlot: slot, sourceIdx: idx })
                              e.dataTransfer.setData('text/plain', payload)
                              e.dataTransfer.effectAllowed = 'move'
                              setDragInfo({ sourceSlot: slot, sourceIdx: idx, teacherId: assignment.teacherId, studentIds: [...assignment.studentIds] })
                            }}
                            onDragOver={(e) => {
                              // Accept student drops onto this assignment block
                              if (isStudentDrag && !assignment.isRegular && assignment.studentIds.length < 2) {
                                e.preventDefault()
                                e.stopPropagation()
                                e.dataTransfer.dropEffect = 'move'
                              }
                            }}
                            onDrop={(e) => {
                              if (!isStudentDrag) return // let slot-card handle assignment drops
                              e.preventDefault()
                              e.stopPropagation()
                              if (assignment.isRegular || assignment.studentIds.length >= 2) {
                                setDragInfo(null)
                                return
                              }
                              try {
                                const raw = e.dataTransfer.getData('text/plain')
                                if (!raw) { setDragInfo(null); return }
                                const payload = JSON.parse(raw) as { sourceSlot: string; sourceIdx: number; studentDragId?: string }
                                if (payload.studentDragId) {
                                  // Prevent dropping on the same assignment
                                  if (payload.sourceSlot === slot && payload.sourceIdx === idx) {
                                    setDragInfo(null)
                                    return
                                  }
                                  void moveStudentToSlot(payload.sourceSlot, payload.sourceIdx, payload.studentDragId, slot, idx)
                                }
                              } catch { /* ignore */ }
                              setDragInfo(null)
                            }}
                            onDragEnd={() => setDragInfo(null)}
                            style={{ position: 'relative' }}
                          >
                            {!assignment.isRegular && (
                              <button
                                type="button"
                                className="pair-delete-btn"
                                title="このペアを削除"
                                onClick={() => void setSlotTeacher(slot, idx, '')}
                              >
                                ×
                              </button>
                            )}
                            {assignment.isRegular && <span className="badge regular-badge" title="通常授業">★</span>}
                            {isIncompatiblePair && <span className="badge incompatible-badge" title="制約不可">⚠</span>}
                            {isAutoAdded && <span className="badge auto-diff-badge auto-diff-badge-new" title={changeDetail || '自動提案で新規追加'}>NEW</span>}
                            {isAutoChanged && <span className="badge auto-diff-badge auto-diff-badge-update" title={changeDetail || '自動提案で再割当'}>UPDATE</span>}
                            <select
                              value={assignment.teacherId}
                              onChange={(e) => void setSlotTeacher(slot, idx, e.target.value)}
                              disabled={assignment.isRegular}
                            >
                              <option value="">{instructorLabel}を選択</option>
                              {instructors
                                .filter((inst) => {
                                  // Always show currently assigned instructor
                                  if (inst.id === assignment.teacherId) return true
                                  // Only show instructors who have availability for this slot
                                  return hasAvailability(data.availability, instructorPersonType, inst.id, slot)
                                })
                                .map((inst) => {
                                const usedElsewhere = usedTeacherIds.has(inst.id) && inst.id !== assignment.teacherId
                                return (
                                  <option key={inst.id} value={inst.id} disabled={usedElsewhere}>
                                    {inst.name}{usedElsewhere ? ' (割当済)' : ''}
                                  </option>
                                )
                              })}
                            </select>

                            {assignment.teacherId && (
                              <>
                                {(isMendan ? [0] : [0, 1]).map((pos) => {
                                  const otherStudentId = assignment.studentIds[pos === 0 ? 1 : 0] ?? ''
                                  const currentStudentId = assignment.studentIds[pos] ?? ''
                                  const currentStudent = data.students.find((s) => s.id === currentStudentId)
                                  const studentSubject = currentStudentId
                                    ? getStudentSubject(assignment, currentStudentId)
                                    : ''
                                  // Subject options for THIS student: teacher can teach AND student can learn
                                  const studentSubjectOptions = selectedTeacher && currentStudent
                                    ? selectedTeacher.subjects.filter((subj) => currentStudent.subjects.includes(subj))
                                    : (selectedTeacher?.subjects ?? [])
                                  // For student dropdown filter: any teacher subject the student can learn
                                  const teacherSubjects = selectedTeacher?.subjects ?? []
                                  return (
                                    <div key={pos} className={`student-select-row${currentStudentId && !assignment.isRegular ? ' student-draggable' : ''}`}
                                      draggable={!!currentStudentId && !assignment.isRegular}
                                      onDragStart={(e) => {
                                        if (!currentStudentId || assignment.isRegular) { e.preventDefault(); return }
                                        e.stopPropagation()
                                        const payload = JSON.stringify({ sourceSlot: slot, sourceIdx: idx, studentDragId: currentStudentId })
                                        e.dataTransfer.setData('text/plain', payload)
                                        e.dataTransfer.effectAllowed = 'move'
                                        setDragInfo({
                                          sourceSlot: slot,
                                          sourceIdx: idx,
                                          teacherId: '',
                                          studentIds: [currentStudentId],
                                          studentDragId: currentStudentId,
                                          studentDragSubject: studentSubject,
                                        })
                                      }}
                                      onDragEnd={() => setDragInfo(null)}
                                    >
                                    {currentStudentId && !assignment.isRegular && (
                                      <span className="student-drag-handle" title="ドラッグで移動">⠿</span>
                                    )}                                    <select
                                      value={currentStudentId}
                                      disabled={assignment.isRegular}
                                      onChange={(e) => {
                                        const selectedId = e.target.value
                                        if (selectedId && !isMendan) {
                                          const student = data.students.find((s) => s.id === selectedId)
                                          if (student) {
                                            const pairTag = constraintFor(data.constraints, assignment.teacherId, student.id)
                                            const gradeTag = gradeConstraintFor(data.gradeConstraints ?? [], assignment.teacherId, student.grade, getStudentSubject(assignment, student.id))
                                            const isIncompatible = pairTag === 'incompatible' || gradeTag === 'incompatible'
                                            if (isIncompatible) {
                                              const reasons: string[] = []
                                              if (pairTag === 'incompatible') reasons.push('ペア制約で不可')
                                              if (gradeTag === 'incompatible') reasons.push(`学年制約(${student.grade})で不可`)
                                              const ok = window.confirm(
                                                `⚠️ ${student.name} は制約ルールにより割当不可です。\n理由: ${reasons.join(', ')}\n\nそれでも割り当てますか？`,
                                              )
                                              if (!ok) {
                                                e.target.value = currentStudentId
                                                return
                                              }
                                            }
                                          }
                                        }
                                        void setSlotStudent(slot, idx, pos, selectedId)
                                      }}
                                    >
                                      <option value="">{isMendan ? `保護者を選択` : `生徒${pos + 1}を選択`}</option>
                                      {data.students
                                        .filter((student) => {
                                          // Always show currently assigned student
                                          if (student.id === currentStudentId) return true
                                          if (isMendan) {
                                            // For mendan: must have submitted, and have positive availability for this slot
                                            if (!student.submittedAt) return false
                                            return isParentAvailableForMendan(data.availability, student.id, slot)
                                          }
                                          // Unsubmitted students are unavailable
                                          if (!student.submittedAt) return false
                                          // Filter out students unavailable for this specific slot
                                          if (!isStudentAvailable(student, slot)) return false
                                          // Subject compatibility: student must share at least one subject with the teacher
                                          if (teacherSubjects.length > 0 && !teacherSubjects.some((subj) => student.subjects.includes(subj))) return false
                                          // Remaining slots for at least one teacher-compatible subject must be > 0
                                          const hasRemainingSlots = teacherSubjects.some((subj) => {
                                            if (!student.subjects.includes(subj)) return false
                                            const desired = student.subjectSlots[subj] ?? 0
                                            const assigned = countStudentSubjectLoad(data.assignments, student.id, subj)
                                            const currentlySelected = currentStudentId === student.id
                                            const adjustedAssigned = currentlySelected ? Math.max(0, assigned - 1) : assigned
                                            return desired - adjustedAssigned > 0
                                          })
                                          return hasRemainingSlots
                                        })
                                        .map((student) => {
                                        const pairTag = isMendan ? null : constraintFor(data.constraints, assignment.teacherId, student.id)
                                        const gradeTag = isMendan ? null : gradeConstraintFor(data.gradeConstraints ?? [], assignment.teacherId, student.grade, getStudentSubject(assignment, student.id))
                                        const isIncompatible = pairTag === 'incompatible' || gradeTag === 'incompatible'
                                        const usedInOther = slotAssignments.some(
                                          (a, i) => i !== idx && a.studentIds.includes(student.id),
                                        )
                                        const isSelectedInOtherPosition = student.id === otherStudentId
                                        const disabled = usedInOther || isSelectedInOtherPosition
                                        const tagLabel = isIncompatible ? ' ⚠不可' : ''
                                        const statusLabel = usedInOther ? ' (他ペア)' : ''

                                        return (
                                          <option key={student.id} value={student.id} disabled={disabled}>
                                            {student.name}{tagLabel}{statusLabel}
                                          </option>
                                        )
                                      })}
                                    </select>
                                    {currentStudentId && !assignment.isRegular && !isMendan && (
                                      <select
                                        className="subject-select-inline"
                                        value={studentSubject}
                                        onChange={(e) => void setSlotSubject(slot, idx, e.target.value, currentStudentId)}
                                      >
                                        {studentSubjectOptions.map((subj) => (
                                          <option key={subj} value={subj}>{subj}</option>
                                        ))}
                                      </select>
                                    )}
                                    {currentStudentId && assignment.isRegular && !isMendan && (
                                      <span className="subject-label-inline">{studentSubject}</span>
                                    )}
                                    </div>
                                  )
                                })}
                              </>
                            )}

                          </div>
                        )
                      })}
                      {(() => {
                        const idleTeachers = instructors.filter(
                          (t) =>
                            hasAvailability(data.availability, instructorPersonType, t.id, slot) &&
                            !usedTeacherIds.has(t.id),
                        )
                        if (idleTeachers.length === 0) return null
                        return (
                          <div className="idle-teachers-compact" style={{ marginTop: '6px', fontSize: '0.85em', color: '#888' }}>
                            {idleTeachers.map((t) => (
                              <span key={t.id} style={{ marginRight: '6px', whiteSpace: 'nowrap' }}>
                                {t.name}
                              </span>
                            ))}
                          </div>
                        )
                      })()}
                    </div>
                  </div>
                )
              })}
            </div>
          </div>
        </>
      )}

    </div>
  )
}

// Helper: Get dates in range excluding holidays
const getDatesInRange = (settings: SessionData['settings']): string[] => {
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

// Teacher Input Component
// Teacher/Manager Input Component
const TeacherInputPage = ({
  sessionId,
  data,
  teacher,
  returnToAdminOnComplete,
  personKeyPrefix = 'teacher',
}: {
  sessionId: string
  data: SessionData
  teacher: Teacher
  returnToAdminOnComplete: boolean
  personKeyPrefix?: 'teacher' | 'manager'
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const showDevRandom = true

  // Find regular lesson slots for this teacher (date_slotNum keys)
  const regularSlotKeys = useMemo(() => {
    const keys = new Set<string>()
    const teacherLessons = data.regularLessons.filter((l) => l.teacherId === teacher.id)
    for (const date of dates) {
      const dayOfWeek = getIsoDayOfWeek(date)
      for (const lesson of teacherLessons) {
        if (lesson.dayOfWeek === dayOfWeek) {
          keys.add(`${date}_${lesson.slotNumber}`)
        }
      }
    }
    return keys
  }, [dates, data.regularLessons, teacher.id])

  const [localAvailability, setLocalAvailability] = useState<Set<string>>(() => {
    const key = personKey(personKeyPrefix, teacher.id)
    const saved = new Set(data.availability[key] ?? [])
    // Include regular lesson slots as forced available
    for (const rk of regularSlotKeys) saved.add(rk)
    return saved
  })
  const toggleSlot = (date: string, slotNum: number) => {
    const slotKey = `${date}_${slotNum}`
    // Cannot toggle regular lesson slots
    if (regularSlotKeys.has(slotKey)) return
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      if (next.has(slotKey)) {
        next.delete(slotKey)
      } else {
        next.add(slotKey)
      }
      return next
    })
  }

  const toggleDateAllSlots = (date: string) => {
    const allSlotKeys = Array.from({ length: data.settings.slotsPerDay }, (_, i) => `${date}_${i + 1}`)
    const nonRegularKeys = allSlotKeys.filter((sk) => !regularSlotKeys.has(sk))
    if (nonRegularKeys.length === 0) return
    const allOn = nonRegularKeys.every((sk) => localAvailability.has(sk))
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      if (allOn) {
        for (const sk of nonRegularKeys) next.delete(sk)
      } else {
        for (const sk of nonRegularKeys) next.add(sk)
      }
      return next
    })
  }

  const toggleColumnAllSlots = (slotNum: number) => {
    const targetKeys = dates
      .map((date) => `${date}_${slotNum}`)
      .filter((slotKey) => !regularSlotKeys.has(slotKey))
    if (targetKeys.length === 0) return
    const allOn = targetKeys.every((slotKey) => localAvailability.has(slotKey))
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      if (allOn) {
        for (const slotKey of targetKeys) next.delete(slotKey)
      } else {
        for (const slotKey of targetKeys) next.add(slotKey)
      }
      return next
    })
  }

  const handleSubmit = () => {
    const key = personKey(personKeyPrefix, teacher.id)
    // Ensure regular lesson slots are always included
    const merged = new Set(localAvailability)
    for (const rk of regularSlotKeys) merged.add(rk)
    const availabilityArray = Array.from(merged)

    // Determine if this is initial or update submission
    const isUpdate = !!(data.teacherSubmittedAt?.[teacher.id])
    const logEntry: SubmissionLogEntry = {
      personId: teacher.id,
      personType: personKeyPrefix === 'manager' ? 'teacher' : 'teacher',
      submittedAt: Date.now(),
      type: isUpdate ? 'update' : 'initial',
      availability: availabilityArray,
    }

    const next: SessionData = {
      ...data,
      availability: {
        ...data.availability,
        [key]: availabilityArray,
      },
      teacherSubmittedAt: {
        ...(data.teacherSubmittedAt ?? {}),
        [teacher.id]: Date.now(),
      },
      submissionLog: [...(data.submissionLog ?? []), logEntry],
    }
    saveSession(sessionId, next).catch(() => { /* ignore */ })
    const submittedAt = new Date().toLocaleString('ja-JP')
    const availCount = availabilityArray.length
    const personTypeLabel = personKeyPrefix === 'manager' ? '講師' : '講師'
    void downloadSubmissionReceiptPdf({
      sessionName: data.settings.name,
      personName: teacher.name,
      personType: personTypeLabel,
      submittedAt,
      details: [`出勤可能コマ数: ${availCount}コマ`],
      isUpdate,
    }).finally(() => {
      navigate(`/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
    })
  }

  // --- Manager-specific: per-day time range input ---
  const isManagerMode = personKeyPrefix === 'manager'
  const mStartHour = data.settings.mendanStartHour ?? 10
  const mEndHour = mStartHour + data.settings.slotsPerDay // exclusive end

  // Compute per-day start/end from current availability
  const getRange = (date: string): { start: number; end: number } | null => {
    let minSlot = Infinity
    let maxSlot = -Infinity
    for (let s = 1; s <= data.settings.slotsPerDay; s++) {
      if (localAvailability.has(`${date}_${s}`)) {
        if (s < minSlot) minSlot = s
        if (s > maxSlot) maxSlot = s
      }
    }
    if (minSlot === Infinity) return null
    return { start: mStartHour - 1 + minSlot, end: mStartHour - 1 + maxSlot + 1 }
  }

  const setRange = (date: string, startHour: number, endHour: number) => {
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      // Clear all slots for this date then set range
      for (let s = 1; s <= data.settings.slotsPerDay; s++) {
        next.delete(`${date}_${s}`)
      }
      const slotStart = startHour - mStartHour + 1
      const slotEnd = endHour - mStartHour + 1
      for (let s = slotStart; s < slotEnd; s++) {
        if (s >= 1 && s <= data.settings.slotsPerDay) {
          next.add(`${date}_${s}`)
        }
      }
      return next
    })
  }

  const clearRange = (date: string) => {
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      for (let s = 1; s <= data.settings.slotsPerDay; s++) {
        next.delete(`${date}_${s}`)
      }
      return next
    })
  }

  // eslint-disable-next-line react-hooks/rules-of-hooks
  const [defaultStart, setDefaultStart] = useState(13)
  // eslint-disable-next-line react-hooks/rules-of-hooks
  const [defaultEnd, setDefaultEnd] = useState(17)

  const applyDefaultToAll = () => {
    setLocalAvailability(() => {
      const next = new Set<string>()
      for (const rk of regularSlotKeys) next.add(rk)
      const slotStart = defaultStart - mStartHour + 1
      const slotEnd = defaultEnd - mStartHour + 1
      for (const date of dates) {
        for (let s = slotStart; s < slotEnd; s++) {
          if (s >= 1 && s <= data.settings.slotsPerDay) {
            next.add(`${date}_${s}`)
          }
        }
      }
      return next
    })
  }

  // Count total available slots
  const totalAvailableSlots = (() => {
    let count = 0
    for (const date of dates) {
      for (let s = 1; s <= data.settings.slotsPerDay; s++) {
        if (localAvailability.has(`${date}_${s}`)) count++
      }
    }
    return count
  })()

  if (isManagerMode) {
    return (
      <div className="availability-container">
        <div className="availability-header">
          <h2>{data.settings.name} - マネージャー面談可能時間入力</h2>
          <p>対象: <strong>{teacher.name}</strong></p>
          <p className="muted">日ごとに面談可能な時間帯を設定してください。</p>
        </div>

        {/* Default time range + apply all */}
        <div className="panel" style={{ marginBottom: '12px', padding: '12px 16px' }}>
          <div className="row" style={{ gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
            <span style={{ fontSize: '14px', fontWeight: 600 }}>一括設定:</span>
            <select value={defaultStart} onChange={(e) => setDefaultStart(Number(e.target.value))} style={{ fontSize: '14px', padding: '4px 8px' }}>
              {Array.from({ length: mEndHour - mStartHour }, (_, i) => mStartHour + i).map((h) => (
                <option key={h} value={h}>{h}:00</option>
              ))}
            </select>
            <span>〜</span>
            <select value={defaultEnd} onChange={(e) => setDefaultEnd(Number(e.target.value))} style={{ fontSize: '14px', padding: '4px 8px' }}>
              {Array.from({ length: mEndHour - mStartHour }, (_, i) => mStartHour + i + 1).map((h) => (
                <option key={h} value={h}>{h}:00</option>
              ))}
            </select>
            <button className="btn" type="button" onClick={applyDefaultToAll} style={{ fontSize: '13px' }}>
              全日に適用
            </button>
          </div>
        </div>

        {/* Per-day time range rows */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
          {dates.map((date) => {
            const range = getRange(date)
            const dayLabel = formatShortDate(date)
            return (
              <div
                key={date}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '8px 12px',
                  background: range ? '#f0f9ff' : '#f9fafb',
                  borderRadius: '8px',
                  border: range ? '1px solid #bae6fd' : '1px solid #e5e7eb',
                  flexWrap: 'wrap',
                }}
              >
                <span style={{ minWidth: '75px', fontWeight: 600, fontSize: '14px' }}>{dayLabel}</span>
                <select
                  value={range?.start ?? defaultStart}
                  onChange={(e) => {
                    const newStart = Number(e.target.value)
                    const currentEnd = range?.end ?? defaultEnd
                    const end = Math.max(newStart + 1, currentEnd)
                    setRange(date, newStart, end)
                  }}
                  style={{ fontSize: '14px', padding: '4px 6px' }}
                >
                  {Array.from({ length: mEndHour - mStartHour }, (_, i) => mStartHour + i).map((h) => (
                    <option key={h} value={h}>{h}:00</option>
                  ))}
                </select>
                <span style={{ color: '#6b7280' }}>〜</span>
                <select
                  value={range?.end ?? defaultEnd}
                  onChange={(e) => {
                    const newEnd = Number(e.target.value)
                    const currentStart = range?.start ?? defaultStart
                    const start = Math.min(currentStart, newEnd - 1)
                    setRange(date, start, newEnd)
                  }}
                  style={{ fontSize: '14px', padding: '4px 6px' }}
                >
                  {Array.from({ length: mEndHour - mStartHour }, (_, i) => mStartHour + i + 1).map((h) => (
                    <option key={h} value={h}>{h}:00</option>
                  ))}
                </select>
                {/* Visual time bar */}
                <div style={{ display: 'flex', gap: '2px', marginLeft: '4px' }}>
                  {Array.from({ length: data.settings.slotsPerDay }, (_, i) => {
                    const slotKey = `${date}_${i + 1}`
                    const isOn = localAvailability.has(slotKey)
                    return (
                      <div
                        key={i}
                        title={`${mStartHour + i}:00`}
                        style={{
                          width: '14px',
                          height: '20px',
                          borderRadius: '3px',
                          background: isOn ? '#3b82f6' : '#e5e7eb',
                          cursor: 'pointer',
                          transition: 'background 0.15s',
                        }}
                        onClick={() => toggleSlot(date, i + 1)}
                      />
                    )
                  })}
                </div>
                {range ? (
                  <button
                    className="btn secondary"
                    type="button"
                    style={{ fontSize: '12px', padding: '2px 8px', marginLeft: 'auto' }}
                    onClick={() => clearRange(date)}
                  >
                    クリア
                  </button>
                ) : (
                  <span style={{ fontSize: '12px', color: '#9ca3af', marginLeft: 'auto' }}>未設定</span>
                )}
              </div>
            )
          })}
        </div>

        <div className="submit-section" style={{ marginTop: '16px' }}>
          <p className="muted" style={{ fontSize: '13px', marginBottom: '8px' }}>
            設定済み: {dates.filter((d) => getRange(d) !== null).length} / {dates.length} 日　（合計 {totalAvailableSlots} コマ）
          </p>
          {showDevRandom && (
            <button
              className="btn secondary"
              type="button"
              style={{ marginBottom: '8px', fontSize: '0.85em' }}
              onClick={() => {
                const next = new Set<string>()
                for (const date of dates) {
                  const randStart = mStartHour + Math.floor(Math.random() * 4) + 1
                  const randEnd = randStart + Math.floor(Math.random() * 5) + 2
                  const slotStart = randStart - mStartHour + 1
                  const slotEnd = Math.min(randEnd - mStartHour + 1, data.settings.slotsPerDay + 1)
                  for (let s = slotStart; s < slotEnd; s++) {
                    next.add(`${date}_${s}`)
                  }
                }
                setLocalAvailability(next)
              }}
            >
              🎲 ランダム入力 (DEV)
            </button>
          )}
          <button
            className="submit-btn"
            onClick={handleSubmit}
            type="button"
            disabled={totalAvailableSlots === 0}
          >
            送信
          </button>
        </div>
      </div>
    )
  }

  return (
    <div className="availability-container">
      <div className="availability-header">
        <h2>{data.settings.name} - 講師希望入力</h2>
        <p>
          対象: <strong>{teacher.name}</strong>
        </p>
        <p className="muted">出席可能なコマをタップして選択してください。「通」は通常授業（変更不可）です。</p>
      </div>

      <div className="teacher-table-wrapper">
        <table className="teacher-table compact-grid">
          <thead>
            <tr>
              <th className="date-header">日付</th>
              {Array.from({ length: data.settings.slotsPerDay }, (_, i) => (
                <th
                  key={i}
                  style={{ cursor: 'pointer', userSelect: 'none' }}
                  onClick={() => toggleColumnAllSlots(i + 1)}
                  title="この時限を一括切替"
                >
                  {`${i + 1}限`}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {dates.map((date) => {
              const nonRegularKeys = Array.from({ length: data.settings.slotsPerDay }, (_, i) => `${date}_${i + 1}`).filter((sk) => !regularSlotKeys.has(sk))
              const allOn = nonRegularKeys.length > 0 && nonRegularKeys.every((sk) => localAvailability.has(sk))
              return (
              <tr key={date}>
                <td
                  className="date-cell"
                  style={{ cursor: 'pointer', userSelect: 'none' }}
                  onClick={() => toggleDateAllSlots(date)}
                  title="全時限を一括切替"
                >
                  <span style={{ fontWeight: allOn ? 700 : 400, color: allOn ? '#2563eb' : undefined }}>
                    {formatShortDate(date)}
                  </span>
                </td>
                {Array.from({ length: data.settings.slotsPerDay }, (_, i) => {
                  const slotNum = i + 1
                  const slotKey = `${date}_${slotNum}`
                  const isRegular = regularSlotKeys.has(slotKey)
                  const isOn = localAvailability.has(slotKey)
                  return (
                    <td key={slotNum}>
                      <button
                        className={`teacher-slot-btn ${isRegular ? 'regular' : isOn ? 'active' : ''}`}
                        onClick={() => toggleSlot(date, slotNum)}
                        type="button"
                        disabled={isRegular}
                      >
                        {isRegular ? '通' : isOn ? '○' : ''}
                      </button>
                    </td>
                  )
                })}
              </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      <div className="submit-section">
        {showDevRandom && (
          <button
            className="btn secondary"
            type="button"
            style={{ marginBottom: '8px', fontSize: '0.85em' }}
            onClick={() => {
              // Randomly toggle ~60% of non-regular slots as available
              const next = new Set(regularSlotKeys)
              for (const date of dates) {
                for (let s = 1; s <= data.settings.slotsPerDay; s++) {
                  const sk = `${date}_${s}`
                  if (!regularSlotKeys.has(sk) && Math.random() < 0.6) next.add(sk)
                }
              }
              setLocalAvailability(next)
            }}
          >
            🎲 ランダム入力 (DEV)
          </button>
        )}
        <button
          className="submit-btn"
          onClick={handleSubmit}
          type="button"
        >
          送信
        </button>
      </div>
    </div>
  )
}

// Student Input Component
const StudentInputPage = ({
  sessionId,
  data,
  student,
  returnToAdminOnComplete,
}: {
  sessionId: string
  data: SessionData
  student: Student
  returnToAdminOnComplete: boolean
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const showDevRandom = true
  const [subjectSlots, setSubjectSlots] = useState<Record<string, number>>(
    student.subjectSlots ?? {},
  )

  // Initialize unavailable slots from existing data (migrate from legacy unavailableDates if needed)
  const [unavailableSlots, setUnavailableSlots] = useState<Set<string>>(() => {
    const initial = new Set(student.unavailableSlots ?? [])
    // Migrate legacy: if unavailableDates exist and unavailableSlots is empty, expand dates to all slots
    if (initial.size === 0 && (student.unavailableDates ?? []).length > 0) {
      for (const date of student.unavailableDates) {
        for (let s = 1; s <= data.settings.slotsPerDay; s++) {
          initial.add(`${date}_${s}`)
        }
      }
    }
    return initial
  })

  const regularSlotKeys = useMemo(() => {
    const keys = new Set<string>()
    const studentLessons = data.regularLessons.filter((lesson) => lesson.studentIds.includes(student.id))
    for (const date of dates) {
      const dayOfWeek = getIsoDayOfWeek(date)
      for (const lesson of studentLessons) {
        if (lesson.dayOfWeek === dayOfWeek) {
          keys.add(`${date}_${lesson.slotNumber}`)
        }
      }
    }
    return keys
  }, [dates, data.regularLessons, student.id])

  const toggleSlot = (slotKey: string) => {
    // Check if this slot has a regular lesson
    const [date, slotNumStr] = slotKey.split('_')
    const slotNum = Number(slotNumStr)
    const dow = getIsoDayOfWeek(date)
    const hasRegular = data.regularLessons.some(
      (l) => l.studentIds.includes(student.id) && l.dayOfWeek === dow && l.slotNumber === slotNum,
    )
    if (hasRegular && !unavailableSlots.has(slotKey)) {
      const confirmed = window.confirm(
        `この時限には通常授業がありますが、出席不可としますか？`,
      )
      if (!confirmed) return
    }

    setUnavailableSlots((prev) => {
      const next = new Set(prev)
      if (next.has(slotKey)) {
        next.delete(slotKey)
      } else {
        next.add(slotKey)
      }
      return next
    })
  }

  const toggleDateAllSlots = (date: string) => {
    const nonRegularKeys = Array.from({ length: data.settings.slotsPerDay }, (_, i) => `${date}_${i + 1}`)
      .filter((slotKey) => !regularSlotKeys.has(slotKey))
    if (nonRegularKeys.length === 0) return
    const allMarked = nonRegularKeys.every((slotKey) => unavailableSlots.has(slotKey))
    setUnavailableSlots((prev) => {
      const next = new Set(prev)
      if (allMarked) {
        for (const slotKey of nonRegularKeys) next.delete(slotKey)
      } else {
        for (const slotKey of nonRegularKeys) next.add(slotKey)
      }
      return next
    })
  }

  const toggleColumnAllSlots = (slotNum: number) => {
    const nonRegularKeys = dates
      .map((date) => `${date}_${slotNum}`)
      .filter((slotKey) => !regularSlotKeys.has(slotKey))
    if (nonRegularKeys.length === 0) return
    const allMarked = nonRegularKeys.every((slotKey) => unavailableSlots.has(slotKey))
    setUnavailableSlots((prev) => {
      const next = new Set(prev)
      if (allMarked) {
        for (const slotKey of nonRegularKeys) next.delete(slotKey)
      } else {
        for (const slotKey of nonRegularKeys) next.add(slotKey)
      }
      return next
    })
  }

  const handleSubjectSlotsChange = (subject: string, value: string) => {
    const numValue = Number(value)
    setSubjectSlots((prev) => ({
      ...prev,
      [subject]: Number.isNaN(numValue) || numValue < 0 ? 0 : Math.floor(numValue),
    }))
  }

  const handleSubmit = () => {
    const subjects = Object.entries(subjectSlots)
      .filter(([, count]) => count > 0)
      .map(([subject]) => subject)
    // Derive unavailableDates from unavailableSlots (dates where ALL slots are unavailable)
    const dateSlotCounts = new Map<string, number>()
    for (const sk of unavailableSlots) {
      const d = sk.split('_')[0]
      dateSlotCounts.set(d, (dateSlotCounts.get(d) ?? 0) + 1)
    }
    const derivedUnavailDates = [...dateSlotCounts.entries()]
      .filter(([, count]) => count >= data.settings.slotsPerDay)
      .map(([d]) => d)

    // Determine if this is initial or update submission
    const isUpdate = !!(student.submittedAt)
    const logEntry: SubmissionLogEntry = {
      personId: student.id,
      personType: 'student',
      submittedAt: Date.now(),
      type: isUpdate ? 'update' : 'initial',
      subjects,
      subjectSlots,
      unavailableSlots: Array.from(unavailableSlots),
    }

    const updatedStudents = data.students.map((s) =>
      s.id === student.id
        ? {
            ...s,
            subjects,
            subjectSlots,
            unavailableDates: derivedUnavailDates,
            preferredSlots: [],
            unavailableSlots: Array.from(unavailableSlots),
            submittedAt: Date.now(),
          }
        : s,
    )

    const next: SessionData = {
      ...data,
      students: updatedStudents,
      submissionLog: [...(data.submissionLog ?? []), logEntry],
    }
    saveSession(sessionId, next).catch(() => { /* ignore */ })
    const submittedAt = new Date().toLocaleString('ja-JP')
    const subjectDetails = Object.entries(subjectSlots)
      .filter(([, count]) => count > 0)
      .map(([subj, count]) => `${subj}: ${count}コマ`)
    const unavailCount = unavailableSlots.size
    void downloadSubmissionReceiptPdf({
      sessionName: data.settings.name,
      personName: student.name,
      personType: '生徒',
      submittedAt,
      details: [...subjectDetails, `不可コマ数: ${unavailCount}`],
      isUpdate,
    }).finally(() => {
      navigate(`/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
    })
  }

  return (    <div className="availability-container">
      <div className="availability-header">
        <h2>{data.settings.name} - 生徒希望入力</h2>
        <p>
          対象: <strong>{student.name}</strong>
        </p>
      </div>

      <div className="student-form-section">
        <h3>希望科目・コマ数</h3>
        <p className="muted">受講を希望する科目を追加し、コマ数を入力してください。</p>
        <div className="subject-slot-entries">
          {Object.entries(subjectSlots)
            .filter(([, count]) => count > 0 || FIXED_SUBJECTS.includes(Object.keys(subjectSlots).find((k) => subjectSlots[k] === count) ?? ''))
            .filter(([subj]) => FIXED_SUBJECTS.includes(subj))
            .map(([subject, count]) => (
              <div key={subject} className="subject-slot-entry">
                <span style={{ fontWeight: 600, minWidth: '28px' }}>{subject}</span>
                <input
                  type="number"
                  min="1"
                  value={count || ''}
                  onChange={(e) => handleSubjectSlotsChange(subject, e.target.value)}
                  placeholder="コマ数"
                />
                <span className="form-unit">コマ</span>
                <button
                  className="subject-slot-remove"
                  type="button"
                  onClick={() => {
                    setSubjectSlots((prev) => {
                      const next = { ...prev }
                      delete next[subject]
                      return next
                    })
                  }}
                >
                  ×
                </button>
              </div>
            ))}
        </div>
        {(() => {
          const selectedSubjects = Object.keys(subjectSlots).filter((s) => FIXED_SUBJECTS.includes(s))
          const availableSubjects = FIXED_SUBJECTS.filter((s) => !selectedSubjects.includes(s))
          if (availableSubjects.length === 0) return null
          return (
            <div style={{ marginTop: '12px' }}>
              <select
                value=""
                onChange={(e) => {
                  const v = e.target.value
                  if (v) {
                    setSubjectSlots((prev) => ({ ...prev, [v]: 1 }))
                  }
                }}
                style={{ padding: '8px', borderRadius: '8px', border: '1px solid #d1d5db', fontSize: '16px' }}
              >
                <option value="">＋ 科目を追加</option>
                {availableSubjects.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </div>
          )
        })()}
      </div>

      <div className="student-form-section">
        <h3>出席不可コマ</h3>
        <p className="muted">出席できないコマをタップして選択してください。日付をタップすると全時限を一括切替できます。</p>
        <div className="teacher-table-wrapper student-no-scroll">
          <table className="teacher-table compact-grid">
            <thead>
              <tr>
                <th className="date-header">日付</th>
                {Array.from({ length: data.settings.slotsPerDay }, (_, i) => (
                  <th
                    key={i}
                    style={{ cursor: 'pointer', userSelect: 'none' }}
                    onClick={() => toggleColumnAllSlots(i + 1)}
                    title="この時限を一括切替"
                  >
                    {`${i + 1}限`}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {dates.map((date) => {
                const allSlotKeys = Array.from({ length: data.settings.slotsPerDay }, (_, i) => `${date}_${i + 1}`)
                const nonRegularKeys = allSlotKeys.filter((slotKey) => !regularSlotKeys.has(slotKey))
                const allMarked = nonRegularKeys.length > 0 && nonRegularKeys.every((slotKey) => unavailableSlots.has(slotKey))
                return (
                  <tr key={date}>
                    <td
                      className="date-cell"
                      style={{ cursor: 'pointer', userSelect: 'none' }}
                      onClick={() => toggleDateAllSlots(date)}
                      title="全時限を一括切替"
                    >
                      <span style={{ fontWeight: allMarked ? 700 : 400, color: allMarked ? '#dc2626' : undefined }}>
                        {formatShortDate(date)}
                      </span>
                    </td>
                    {Array.from({ length: data.settings.slotsPerDay }, (_, i) => {
                      const slotNum = i + 1
                      const slotKey = `${date}_${slotNum}`
                      const isUnavail = unavailableSlots.has(slotKey)
                      const hasRegular = regularSlotKeys.has(slotKey)
                      return (
                        <td key={slotNum}>
                          <button
                            className={`teacher-slot-btn ${isUnavail ? 'unavail' : ''} ${hasRegular && !isUnavail ? 'regular' : ''}`}
                            onClick={() => toggleSlot(slotKey)}
                            type="button"
                          >
                            {isUnavail ? '✕' : hasRegular ? '通' : ''}
                          </button>
                        </td>
                      )
                    })}
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div className="submit-section">
        {showDevRandom && (
          <button
            className="btn secondary"
            type="button"
            style={{ marginBottom: '8px', fontSize: '0.85em' }}
            onClick={() => {
              // Randomly pick 1-3 subjects with 1-4 slots each
              const shuffled = [...FIXED_SUBJECTS].sort(() => Math.random() - 0.5)
              const count = 1 + Math.floor(Math.random() * 3)
              const randomSubjects: Record<string, number> = {}
              for (let i = 0; i < count && i < shuffled.length; i++) {
                randomSubjects[shuffled[i]] = 1 + Math.floor(Math.random() * 4)
              }
              setSubjectSlots(randomSubjects)
              // Randomly mark ~20% of slots as unavailable
              const next = new Set<string>()
              for (const date of dates) {
                for (let s = 1; s <= data.settings.slotsPerDay; s++) {
                  if (Math.random() < 0.2) next.add(`${date}_${s}`)
                }
              }
              setUnavailableSlots(next)
            }}
          >
            🎲 ランダム入力 (DEV)
          </button>
        )}
        {(() => {
          const totalDesired = Object.values(subjectSlots).reduce((s, c) => s + c, 0)
          const totalSlots = dates.length * data.settings.slotsPerDay
          const totalAvailable = totalSlots - unavailableSlots.size
          const noSubjects = Object.entries(subjectSlots).filter(([, c]) => c > 0).length === 0
          const overAvailable = totalDesired > totalAvailable
          return (
            <>
              {noSubjects && (
                <p style={{ color: '#dc2626', fontWeight: 600, marginBottom: '8px', fontSize: '14px' }}>※ 科目を1つ以上選択してください</p>
              )}
              {!noSubjects && overAvailable && (
                <p style={{ color: '#dc2626', fontWeight: 600, marginBottom: '8px', fontSize: '14px' }}>
                  ※ 希望コマ数({totalDesired})が出席可能コマ数({totalAvailable})を超えています
                </p>
              )}
              <button
                className="submit-btn"
                onClick={handleSubmit}
                type="button"
                disabled={noSubjects || overAvailable}
              >
                送信
              </button>
            </>
          )
        })()}
      </div>
    </div>
  )
}

// Mendan Parent Input Component — parents mark available time slots
const MendanParentInputPage = ({
  sessionId,
  data,
  student,
  returnToAdminOnComplete,
}: {
  sessionId: string
  data: SessionData
  student: Student
  returnToAdminOnComplete: boolean
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const showDevRandom = true

  // Compute which slots have at least one manager available
  const managerAvailableSlots = useMemo(() => {
    const slots = new Set<string>()
    for (const manager of (data.managers ?? [])) {
      const key = personKey('manager', manager.id)
      for (const sk of (data.availability[key] ?? [])) {
        slots.add(sk)
      }
    }
    return slots
  }, [data])

  const [localAvailability, setLocalAvailability] = useState<Set<string>>(() => {
    const key = personKey('student', student.id)
    return new Set(data.availability[key] ?? [])
  })

  const toggleSlot = (date: string, slotNum: number) => {
    const slotKey = `${date}_${slotNum}`
    // Only allow toggling on slots where at least one manager is available
    if (!managerAvailableSlots.has(slotKey)) return
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      if (next.has(slotKey)) {
        next.delete(slotKey)
      } else {
        next.add(slotKey)
      }
      return next
    })
  }

  const toggleDateAllSlots = (date: string) => {
    const allSlotKeys = Array.from({ length: data.settings.slotsPerDay }, (_, i) => `${date}_${i + 1}`)
    const selectableKeys = allSlotKeys.filter((sk) => managerAvailableSlots.has(sk))
    if (selectableKeys.length === 0) return
    const allOn = selectableKeys.every((sk) => localAvailability.has(sk))
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      if (allOn) {
        for (const sk of selectableKeys) next.delete(sk)
      } else {
        for (const sk of selectableKeys) next.add(sk)
      }
      return next
    })
  }

  const toggleColumnAllSlots = (slotNum: number) => {
    const targetKeys = dates
      .map((date) => `${date}_${slotNum}`)
      .filter((sk) => managerAvailableSlots.has(sk))
    if (targetKeys.length === 0) return
    const allOn = targetKeys.every((sk) => localAvailability.has(sk))
    setLocalAvailability((prev) => {
      const next = new Set(prev)
      if (allOn) {
        for (const sk of targetKeys) next.delete(sk)
      } else {
        for (const sk of targetKeys) next.add(sk)
      }
      return next
    })
  }

  const handleSubmit = () => {
    const key = personKey('student', student.id)
    const availabilityArray = Array.from(localAvailability)

    const logEntry: SubmissionLogEntry = {
      personId: student.id,
      personType: 'student',
      submittedAt: Date.now(),
      type: student.submittedAt ? 'update' : 'initial',
      availability: availabilityArray,
    }

    const updatedStudents = data.students.map((s) =>
      s.id === student.id
        ? { ...s, submittedAt: Date.now() }
        : s,
    )

    const next: SessionData = {
      ...data,
      students: updatedStudents,
      availability: {
        ...data.availability,
        [key]: availabilityArray,
      },
      submissionLog: [...(data.submissionLog ?? []), logEntry],
    }
    saveSession(sessionId, next).catch(() => { /* ignore */ })
    const submittedAt = new Date().toLocaleString('ja-JP')
    const isUpdate = !!(student.submittedAt)
    const availCount = availabilityArray.length
    void downloadSubmissionReceiptPdf({
      sessionName: data.settings.name,
      personName: `${student.name} 保護者`,
      personType: '保護者',
      submittedAt,
      details: [`面談希望コマ数: ${availCount}コマ`],
      isUpdate,
    }).finally(() => {
      navigate(`/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
    })
  }

  // Compute which slot numbers have at least one manager available (for column display)
  const mendanStartHour = data.settings.mendanStartHour ?? 10
  const activeSlotNums = useMemo(() => {
    const nums = new Set<number>()
    for (const sk of managerAvailableSlots) {
      const slotNum = Number(sk.split('_')[1])
      if (!isNaN(slotNum)) nums.add(slotNum)
    }
    return Array.from(nums).sort((a, b) => a - b)
  }, [managerAvailableSlots])

  return (
    <div className="availability-container">
      <div className="availability-header">
        <h2>{data.settings.name} - 保護者面談希望入力</h2>
        <p>
          対象: <strong>{student.name}</strong> 保護者
        </p>
        <p className="muted">面談可能な時間帯をタップして選択してください。色付きのコマはマネージャーが対応可能な時間帯です。</p>
      </div>

      {activeSlotNums.length === 0 ? (
        <div className="panel" style={{ textAlign: 'center', padding: '24px' }}>
          <p style={{ color: '#dc2626', fontWeight: 600 }}>マネージャーの空き時間がまだ登録されていません。</p>
          <p className="muted">マネージャーが面談可能時間を入力すると、ここに選択肢が表示されます。</p>
        </div>
      ) : (
        <div className="teacher-table-wrapper">
          <table className="teacher-table compact-grid">
            <thead>
              <tr>
                <th className="date-header">日付</th>
                {activeSlotNums.map((slotNum) => (
                  <th
                    key={slotNum}
                    style={{ cursor: 'pointer', userSelect: 'none' }}
                    onClick={() => toggleColumnAllSlots(slotNum)}
                    title="この時間帯を一括切替"
                  >
                    {mendanTimeLabel(slotNum, mendanStartHour)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {dates.map((date) => {
                const selectableKeys = activeSlotNums.map((s) => `${date}_${s}`).filter((sk) => managerAvailableSlots.has(sk))
                const allOn = selectableKeys.length > 0 && selectableKeys.every((sk) => localAvailability.has(sk))
                return (
                  <tr key={date}>
                    <td
                      className="date-cell"
                      style={{ cursor: 'pointer', userSelect: 'none' }}
                      onClick={() => toggleDateAllSlots(date)}
                      title="全時限を一括切替"
                    >
                      <span style={{ fontWeight: allOn ? 700 : 400, color: allOn ? '#2563eb' : undefined }}>
                        {formatShortDate(date)}
                      </span>
                    </td>
                    {activeSlotNums.map((slotNum) => {
                      const slotKey = `${date}_${slotNum}`
                      const managerAvail = managerAvailableSlots.has(slotKey)
                      const isOn = localAvailability.has(slotKey)
                      return (
                        <td key={slotNum}>
                          <button
                            className={`teacher-slot-btn ${!managerAvail ? '' : isOn ? 'active' : 'manager-avail'}`}
                            onClick={() => toggleSlot(date, slotNum)}
                            type="button"
                            disabled={!managerAvail}
                            style={!managerAvail ? { opacity: 0.3, cursor: 'not-allowed' } : undefined}
                          >
                            {!managerAvail ? '' : isOn ? '○' : ''}
                          </button>
                        </td>
                      )
                    })}
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}

      <div className="submit-section">
        {showDevRandom && (
          <button
            className="btn secondary"
            type="button"
            style={{ marginBottom: '8px', fontSize: '0.85em' }}
            onClick={() => {
              // Randomly select ~50% of manager-available slots
              const next = new Set<string>()
              for (const sk of managerAvailableSlots) {
                if (Math.random() < 0.5) next.add(sk)
              }
              setLocalAvailability(next)
            }}
          >
            🎲 ランダム入力 (DEV)
          </button>
        )}
        {(() => {
          const selectedCount = localAvailability.size
          return (
            <>
              {selectedCount === 0 && (
                <p style={{ color: '#dc2626', fontWeight: 600, marginBottom: '8px', fontSize: '14px' }}>※ 面談可能な時間帯を1つ以上選択してください</p>
              )}
              <button
                className="submit-btn"
                onClick={handleSubmit}
                type="button"
                disabled={selectedCount === 0}
              >
                送信
              </button>
            </>
          )
        })()}
      </div>
    </div>
  )
}

// ─── Confirmed Calendar View ────────────────────────────────────────
// When the admin has pressed "確定", input URLs show this calendar instead
// of the input form. Content varies by role.
const ConfirmedCalendarView = ({
  data,
  personType,
  personId,
}: {
  data: SessionData
  personType: PersonType
  personId: string
}) => {
  const isMendan = data.settings.sessionType === 'mendan'
  const mendanStart = data.settings.mendanStartHour ?? 10
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const slotsPerDay = data.settings.slotsPerDay

  // Find the person name
  const personName = useMemo(() => {
    if (personType === 'teacher') return data.teachers.find((t) => t.id === personId)?.name ?? ''
    if (personType === 'manager') return (data.managers ?? []).find((m) => m.id === personId)?.name ?? ''
    return data.students.find((s) => s.id === personId)?.name ?? ''
  }, [data, personType, personId])

  // Build a map: date → slot → assignments relevant to this person
  type CalendarCell = { label: string; detail: string; color: string }
  const calendar = useMemo(() => {
    const result: Record<string, Record<number, CalendarCell[]>> = {}
    for (const date of dates) {
      result[date] = {}
      for (let s = 1; s <= slotsPerDay; s++) {
        const slotKey = `${date}_${s}`
        const slotAssignments = data.assignments[slotKey] ?? []
        const cells: CalendarCell[] = []

        if (personType === 'teacher' || personType === 'manager') {
          // Teacher/Manager: show which students and subject
          for (const a of slotAssignments) {
            if (a.teacherId !== personId) continue
            for (const sid of a.studentIds) {
              const student = data.students.find((st) => st.id === sid)
              const subj = getStudentSubject(a, sid)
              if (isMendan) {
                cells.push({
                  label: `${student?.name ?? '?'} 保護者`,
                  detail: personType === 'manager' ? '' : '面談',
                  color: '#dbeafe',
                })
              } else {
                cells.push({
                  label: student?.name ?? '?',
                  detail: subj,
                  color: a.isRegular ? '#dcfce7' : '#fef3c7',
                })
              }
            }
          }
        } else if (personType === 'student') {
          // Student (regular session): show teacher + subject + regular/special
          // Parent (mendan): show when interview is
          for (const a of slotAssignments) {
            if (!a.studentIds.includes(personId)) continue
            if (isMendan) {
              const manager = (data.managers ?? []).find((m) => m.id === a.teacherId)
              cells.push({
                label: '面談',
                detail: manager?.name ? `担当: ${manager.name}` : '',
                color: '#dbeafe',
              })
            } else {
              const teacher = data.teachers.find((t) => t.id === a.teacherId)
              const subj = getStudentSubject(a, personId)
              cells.push({
                label: a.isRegular ? `★ ${subj}` : subj,
                detail: `${teacher?.name ?? '?'}${a.isRegular ? ' (通常)' : ' (特別講習)'}`,
                color: a.isRegular ? '#dcfce7' : '#fef3c7',
              })
            }
          }
        }

        if (cells.length > 0) result[date][s] = cells
      }
    }
    return result
  }, [data, dates, slotsPerDay, personType, personId, isMendan])

  // For mendan mode: only show slots that have any assignment across all dates
  const activeSlotNums = useMemo(() => {
    if (!isMendan) return Array.from({ length: slotsPerDay }, (_, i) => i + 1)
    const nums = new Set<number>()
    for (const date of dates) {
      for (let s = 1; s <= slotsPerDay; s++) {
        const slotKey = `${date}_${s}`
        if ((data.assignments[slotKey] ?? []).length > 0) nums.add(s)
      }
    }
    // Also include manager availability slots for mendan
    if (personType === 'manager' || personType === 'teacher') {
      const pk = `${personType}:${personId}`
      for (const sk of (data.availability[pk] ?? [])) {
        const num = Number(sk.split('_')[1])
        if (!isNaN(num)) nums.add(num)
      }
    }
    return Array.from(nums).sort((a, b) => a - b)
  }, [data, dates, slotsPerDay, isMendan, personType, personId])

  const slotHeader = (slotNum: number): string => {
    if (isMendan) return mendanTimeLabel(slotNum, mendanStart)
    return `${slotNum}限`
  }

  const roleLabel = personType === 'teacher' ? '講師' : personType === 'manager' ? 'マネージャー' : isMendan ? '保護者' : '生徒'

  // Count total assigned slots for this person
  const totalSlots = useMemo(() => {
    let count = 0
    for (const date of dates) {
      for (let s = 1; s <= slotsPerDay; s++) {
        if ((calendar[date]?.[s]?.length ?? 0) > 0) count++
      }
    }
    return count
  }, [calendar, dates, slotsPerDay])

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>{data.settings.name} - スケジュール確認</h2>
        <p><strong>{personName}</strong>（{roleLabel}）のスケジュール</p>
        {isMendan && personType === 'student' && (
          <p className="muted">以下があなたの面談スケジュールです。</p>
        )}
        {!isMendan && personType === 'student' && (
          <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', marginBottom: '8px' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '13px' }}>
              <span style={{ display: 'inline-block', width: '14px', height: '14px', background: '#dcfce7', border: '1px solid #86efac', borderRadius: '3px' }} /> 通常授業
            </span>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '13px' }}>
              <span style={{ display: 'inline-block', width: '14px', height: '14px', background: '#fef3c7', border: '1px solid #fcd34d', borderRadius: '3px' }} /> 特別講習
            </span>
          </div>
        )}
        {!isMendan && personType === 'teacher' && (
          <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', marginBottom: '8px' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '13px' }}>
              <span style={{ display: 'inline-block', width: '14px', height: '14px', background: '#dcfce7', border: '1px solid #86efac', borderRadius: '3px' }} /> 通常授業
            </span>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '13px' }}>
              <span style={{ display: 'inline-block', width: '14px', height: '14px', background: '#fef3c7', border: '1px solid #fcd34d', borderRadius: '3px' }} /> 特別講習
            </span>
          </div>
        )}
        <p className="muted" style={{ fontSize: '13px' }}>合計 {totalSlots} コマ</p>
      </div>

      <div className="panel" style={{ overflowX: 'auto' }}>
        <table className="table" style={{ fontSize: '13px', minWidth: '600px' }}>
          <thead>
            <tr>
              <th style={{ position: 'sticky', left: 0, background: '#f8fafc', zIndex: 2, minWidth: '80px' }}>日付</th>
              {activeSlotNums.map((s) => (
                <th key={s} style={{ textAlign: 'center', minWidth: '100px', whiteSpace: 'nowrap' }}>{slotHeader(s)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {dates.map((date) => {
              const d = new Date(`${date}T00:00:00`)
              const dayName = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()]
              const isWeekend = d.getDay() === 0 || d.getDay() === 6
              return (
                <tr key={date}>
                  <td style={{
                    position: 'sticky', left: 0, background: isWeekend ? '#fef2f2' : '#f8fafc',
                    zIndex: 1, fontWeight: 600, whiteSpace: 'nowrap',
                    color: d.getDay() === 0 ? '#dc2626' : d.getDay() === 6 ? '#2563eb' : undefined,
                  }}>
                    {d.getMonth() + 1}/{d.getDate()}({dayName})
                  </td>
                  {activeSlotNums.map((s) => {
                    const cells = calendar[date]?.[s] ?? []
                    return (
                      <td key={s} style={{ textAlign: 'center', padding: '4px 6px', verticalAlign: 'top' }}>
                        {cells.map((cell, idx) => (
                          <div key={idx} style={{
                            background: cell.color, borderRadius: '4px', padding: '3px 6px',
                            marginBottom: idx < cells.length - 1 ? '2px' : 0,
                            fontSize: '12px', lineHeight: '1.4',
                            whiteSpace: 'nowrap',
                          }}>
                            <span style={{ fontWeight: 600 }}>{cell.label}</span>
                            {cell.detail && <span style={{ fontSize: '11px', color: '#475569', marginLeft: 4 }}>{cell.detail}</span>}
                          </div>
                        ))}
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

const AvailabilityPage = () => {
  const location = useLocation()
  const { sessionId = 'main', personType: rawPersonType = 'teacher', personId: rawPersonId = '' } = useParams()
  const personType = (rawPersonType === 'student' ? 'student' : rawPersonType === 'manager' ? 'manager' : 'teacher') as PersonType
  const personId = useMemo(() => rawPersonId.split(/[?:&]/)[0], [rawPersonId])
  const [data, setData] = useState<SessionData | null>(null)
  const [phase, setPhase] = useState<'loading' | 'ready' | 'not-found' | 'permission-error' | 'timeout'>('loading')
  const [errorDetail, setErrorDetail] = useState('')
  const syncingRef = useRef(false)
  const syncDoneRef = useRef(false)
  const returnToAdminOnComplete = (location.state as { fromAdminInput?: boolean } | null)?.fromAdminInput === true

  useEffect(() => {
    setPhase('loading')
    setData(null)
    setErrorDetail('')
    syncingRef.current = false
    syncDoneRef.current = false

    // Timeout: if nothing loads within 10 seconds, show error
    let timedOut = false
    const timer = setTimeout(() => {
      timedOut = true
      setErrorDetail('読み込みがタイムアウトしました。ネットワーク接続を確認してください。')
      setPhase('timeout')
    }, 10000)

    const unsub = watchSession(
      sessionId,
      (value) => {
        if (timedOut) return
        clearTimeout(timer)

        if (!value) {
          setErrorDetail(`特別講習「${sessionId}」がFirebaseに見つかりません。`)
          setPhase('not-found')
          return
        }

        // Session found — check if the person exists
        const found = personType === 'teacher'
          ? value.teachers.find((t) => t.id === personId)
          : personType === 'manager'
            ? (value.managers ?? []).find((m) => m.id === personId)
            : value.students.find((s) => s.id === personId)

        if (found) {
          setData(value)
          setPhase('ready')
          return
        }

        // Person not found — try syncing master data once
        if (syncDoneRef.current) {
          setData(value)
          setPhase('ready')
          return
        }
        if (syncingRef.current) return

        syncingRef.current = true
        loadMasterData()
          .then((master) => {
            if (!master) {
              syncDoneRef.current = true
              syncingRef.current = false
              setData(value)
              setPhase('ready')
              return
            }
            const mergedStudents = master.students.map((ms) => {
              const existing = value.students.find((s) => s.id === ms.id)
              if (existing) {
                return { ...ms, subjects: existing.subjects, subjectSlots: existing.subjectSlots, unavailableDates: existing.unavailableDates, preferredSlots: existing.preferredSlots ?? [], unavailableSlots: existing.unavailableSlots ?? [], submittedAt: existing.submittedAt }
              }
              return { ...ms, subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], submittedAt: 0 }
            })
            const next: SessionData = {
              ...value,
              teachers: master.teachers,
              students: mergedStudents,
              constraints: master.constraints,
              gradeConstraints: master.gradeConstraints ?? [],
              regularLessons: master.regularLessons,
            }
            setData(next)
            setPhase('ready')
            return saveSession(sessionId, next)
          })
          .then(() => {
            syncDoneRef.current = true
            syncingRef.current = false
          })
          .catch(() => {
            syncDoneRef.current = true
            syncingRef.current = false
            setData((prev) => prev ?? value)
            setPhase('ready')
          })
      },
      (error) => {
        if (timedOut) return
        clearTimeout(timer)
        setErrorDetail(`Firebaseアクセスエラー: ${error.message}`)
        setPhase('permission-error')
      },
    )

    return () => { unsub(); clearTimeout(timer) }
  }, [sessionId, personType, personId])

  const currentPerson = useMemo(() => {
    if (!data) return null
    if (personType === 'teacher') {
      return data.teachers.find((teacher) => teacher.id === personId) ?? null
    }
    if (personType === 'manager') {
      return (data.managers ?? []).find((m) => m.id === personId) ?? null
    }
    return data.students.find((student) => student.id === personId) ?? null
  }, [data, personType, personId])

  if (phase === 'loading') {
    return (
      <div className="app-shell">
        <div className="panel">読み込み中...</div>
      </div>
    )
  }

  if (phase === 'permission-error') {
    return (
      <div className="app-shell">
        <div className="panel">
          <h3 style={{ color: '#dc2626' }}>⚠ Firebaseへのアクセスが拒否されました</h3>
          <p>{errorDetail}</p>
          <p className="muted">
            管理者がFirebase Consoleで以下の設定を行う必要があります：
          </p>
          <ol style={{ fontSize: '13px', lineHeight: 1.8 }}>
            <li><strong>Authentication</strong> → Sign-in method → 「匿名」を有効化</li>
            <li><strong>Firestore Database</strong> → ルール → 以下に書き換え：</li>
          </ol>
          <pre style={{ fontSize: '11px', background: '#f3f4f6', padding: '8px', borderRadius: '4px', whiteSpace: 'pre-wrap' }}>
{`rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if request.auth != null;
    }
  }
}`}
          </pre>
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  if (phase === 'not-found') {
    return (
      <div className="app-shell">
        <div className="panel">
          <h3>特別講習が見つかりません</h3>
          <p>{errorDetail}</p>
          <p className="muted">
            考えられる原因：<br />
            ・管理者がまだ特別講習を作成していない<br />
            ・IDが間違っている<br />
            ・Firestoreへのデータ保存に失敗している
          </p>
          <p className="muted" style={{ fontSize: '11px' }}>
            特別講習ID: {sessionId} / {personType} / {personId}
          </p>
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  if (phase === 'timeout') {
    return (
      <div className="app-shell">
        <div className="panel">
          <h3>タイムアウト</h3>
          <p>{errorDetail}</p>
          <button className="btn" type="button" onClick={() => window.location.reload()}>再読み込み</button>
          <br />
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  if (!data || !currentPerson) {
    return (
      <div className="app-shell">
        <div className="panel">
          入力対象が見つかりません。管理者にURLを確認してください。
          <br />
          <p className="muted" style={{ fontSize: '11px' }}>
            特別講習ID: {sessionId} / {personType} / {personId}
          </p>
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  // When admin has confirmed the schedule, show calendar view instead of input form
  if (data.settings.confirmed) {
    return <ConfirmedCalendarView data={data} personType={personType} personId={personId} />
  }

  if (personType === 'teacher') {
    if ('subjects' in currentPerson && Array.isArray(currentPerson.subjects)) {
      // Check submission period for teachers too
      const now = new Date()
      const startDate = data.settings.submissionStartDate ? new Date(data.settings.submissionStartDate) : null
      const endDate = data.settings.submissionEndDate ? new Date(data.settings.submissionEndDate + 'T23:59:59') : null
      if (startDate && now < startDate) {
        return (
          <div className="app-shell">
            <div className="panel">
              <h3>提出期間前です</h3>
              <p>提出受付開始日: <strong>{data.settings.submissionStartDate}</strong></p>
              <p className="muted">提出期間になるまでお待ちください。</p>
              <Link to="/">ホームに戻る</Link>
            </div>
          </div>
        )
      }
      if (endDate && now > endDate) {
        return (
          <div className="app-shell">
            <div className="panel">
              <h3>提出期間は終了しました</h3>
              <p>提出締切日: <strong>{data.settings.submissionEndDate}</strong></p>
              <p className="muted">期間を過ぎています。管理者にお問い合わせください。</p>
              <Link to="/">ホームに戻る</Link>
            </div>
          </div>
        )
      }
      return <TeacherInputPage sessionId={sessionId} data={data} teacher={currentPerson as Teacher} returnToAdminOnComplete={returnToAdminOnComplete} />
    }
  } else if (personType === 'manager') {
    // Manager availability input — same as teacher but uses 'manager' personKey
    const manager = currentPerson as Manager
    const now = new Date()
    const startDate = data.settings.submissionStartDate ? new Date(data.settings.submissionStartDate) : null
    const endDate = data.settings.submissionEndDate ? new Date(data.settings.submissionEndDate + 'T23:59:59') : null
    if (startDate && now < startDate) {
      return (
        <div className="app-shell">
          <div className="panel">
            <h3>提出期間前です</h3>
            <p>提出受付開始日: <strong>{data.settings.submissionStartDate}</strong></p>
            <p className="muted">提出期間になるまでお待ちください。</p>
            <Link to="/">ホームに戻る</Link>
          </div>
        </div>
      )
    }
    if (endDate && now > endDate) {
      return (
        <div className="app-shell">
          <div className="panel">
            <h3>提出期間は終了しました</h3>
            <p>提出締切日: <strong>{data.settings.submissionEndDate}</strong></p>
            <p className="muted">期間を過ぎています。管理者にお問い合わせください。</p>
            <Link to="/">ホームに戻る</Link>
          </div>
        </div>
      )
    }
    // Wrap manager as a Teacher-like object so TeacherInputPage can be reused
    const managerAsTeacher: Teacher = { id: manager.id, name: manager.name, email: manager.email, subjects: ['面談'], memo: '' }
    return <TeacherInputPage sessionId={sessionId} data={data} teacher={managerAsTeacher} returnToAdminOnComplete={returnToAdminOnComplete} personKeyPrefix="manager" />
  } else if (personType === 'student') {
    if ('grade' in currentPerson && 'subjectSlots' in currentPerson) {
      // Check submission period for students
      const now = new Date()
      const startDate = data.settings.submissionStartDate ? new Date(data.settings.submissionStartDate) : null
      const endDate = data.settings.submissionEndDate ? new Date(data.settings.submissionEndDate + 'T23:59:59') : null
      if (startDate && now < startDate) {
        return (
          <div className="app-shell">
            <div className="panel">
              <h3>提出期間前です</h3>
              <p>提出受付開始日: <strong>{data.settings.submissionStartDate}</strong></p>
              <p className="muted">提出期間になるまでお待ちください。</p>
              <Link to="/">ホームに戻る</Link>
            </div>
          </div>
        )
      }
      if (endDate && now > endDate) {
        return (
          <div className="app-shell">
            <div className="panel">
              <h3>提出期間は終了しました</h3>
              <p>提出締切日: <strong>{data.settings.submissionEndDate}</strong></p>
              <p className="muted">期間を過ぎています。管理者にお問い合わせください。</p>
              <Link to="/">ホームに戻る</Link>
            </div>
          </div>
        )
      }
      return data.settings.sessionType === 'mendan'
        ? <MendanParentInputPage sessionId={sessionId} data={data} student={currentPerson as Student} returnToAdminOnComplete={returnToAdminOnComplete} />
        : <StudentInputPage sessionId={sessionId} data={data} student={currentPerson as Student} returnToAdminOnComplete={returnToAdminOnComplete} />
    }
  }

  return (
    <div className="app-shell">
      <div className="panel">
        入力対象の種別が正しくありません。管理者にURLを確認してください。
        <br />
        <Link to="/">ホームに戻る</Link>
      </div>
    </div>
  )
}

const BootPage = () => {
  const navigate = useNavigate()

  useEffect(() => {
    saveSession('main', createTemplateSession()).catch(() => {})
    navigate('/admin/main', { replace: true })
  }, [navigate])

  return (
    <div className="app-shell">
      <div className="panel">初期化中...</div>
    </div>
  )
}

const CompletionPage = () => {
  const location = useLocation()
  const { sessionId = 'main' } = useParams()
  const returnToAdminOnComplete = (location.state as { returnToAdminOnComplete?: boolean } | null)?.returnToAdminOnComplete === true

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>入力完了</h2>
        <p>データの送信が完了しました。ありがとうございます。</p>
        {returnToAdminOnComplete && (
          <div className="row" style={{ marginTop: '10px' }}>
            <Link className="btn" to={`/admin/${sessionId}`} state={{ skipAuth: true }}>管理画面へ戻る</Link>
          </div>
        )}
      </div>
    </div>
  )
}

function App() {
  const [authReady, setAuthReady] = useState(false)

  useEffect(() => {
    initAuth().finally(() => setAuthReady(true))
  }, [])

  if (!authReady) {
    return (
      <div className="app-shell">
        <div className="panel">接続中...</div>
      </div>
    )
  }

  return (
    <>
      <div className="version-badge">v{APP_VERSION}</div>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/boot" element={<BootPage />} />
        <Route path="/admin/:sessionId" element={<AdminPage />} />
        <Route path="/availability/:sessionId/:personType/:personId" element={<AvailabilityPage />} />
        <Route path="/complete/:sessionId" element={<CompletionPage />} />
      </Routes>
    </>
  )
}

export default App
