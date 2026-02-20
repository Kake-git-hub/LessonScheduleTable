import { useEffect, useMemo, useRef, useState } from 'react'
import { Link, Route, Routes, useLocation, useNavigate, useParams } from 'react-router-dom'
import * as XLSX from 'xlsx'
import './App.css'
import { deleteSession, initAuth, loadMasterData, loadSession, saveAndVerify, saveMasterData, saveSession, watchMasterData, watchSession, watchSessionsList } from './firebase'
import type {
  Assignment,
  ConstraintType,
  GradeConstraint,
  MasterData,
  PairConstraint,
  PersonType,
  RegularLesson,
  SessionData,
  Student,
  SubmissionLogEntry,
  Teacher,
} from './types'
import { buildSlotKeys, formatShortDate, personKey, slotLabel } from './utils/schedule'

const APP_VERSION = '0.4.0'

const GRADE_OPTIONS = ['小4', '小5', '小6', '中1', '中2', '中3', '高1', '高2', '高3']

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
    { id: 't001', name: '田中講師', subjects: ['数', '英'], memo: '数学メイン' },
    { id: 't002', name: '佐藤講師', subjects: ['英', '数'], memo: '英語メイン' },
  ]

  const students: Student[] = [
    {
      id: 's001',
      name: '青木 太郎',
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
): ConstraintType | null => {
  if (!grade) return null
  const hit = gradeConstraints.find((item) => item.teacherId === teacherId && item.grade === grade)
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
    (a) => a.studentIds.includes(studentId) && a.subject === subject && !a.isRegular,
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
    }
  }
  return shortages
}

const assignmentSignature = (assignment: Assignment): string =>
  `${assignment.teacherId}|${assignment.subject}|${[...assignment.studentIds].sort().join('+')}|${assignment.isRegular ? 'R' : 'N'}`

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
): { assignments: Record<string, Assignment[]>; changeLog: ChangeLogEntry[]; changedPairSignatures: Record<string, string[]>; addedPairSignatures: Record<string, string[]> } => {
  const changeLog: ChangeLogEntry[] = []
  const changedPairSigSetBySlot: Record<string, Set<string>> = {}
  const addedPairSigSetBySlot: Record<string, Set<string>> = {}
  const markChangedPair = (slot: string, assignment: Assignment): void => {
    if (assignment.isRegular) return
    if (!hasMeaningfulManualAssignment(assignment)) return
    if (!changedPairSigSetBySlot[slot]) changedPairSigSetBySlot[slot] = new Set<string>()
    changedPairSigSetBySlot[slot].add(assignmentSignature(assignment))
  }
  const markAddedPair = (slot: string, assignment: Assignment): void => {
    if (assignment.isRegular) return
    if (!hasMeaningfulManualAssignment(assignment)) return
    if (!addedPairSigSetBySlot[slot]) addedPairSigSetBySlot[slot] = new Set<string>()
    addedPairSigSetBySlot[slot].add(assignmentSignature(assignment))
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

      // Check teacher still exists
      if (!teacherIds.has(assignment.teacherId)) {
        const usedTeachers = new Set(cleaned.map((a) => a.teacherId))
        const replacement = data.teachers.find((t) => {
          if (usedTeachers.has(t.id)) return false
          if (!hasAvailability(data.availability, 'teacher', t.id, slot)) return false
          return t.subjects.includes(assignment.subject)
        })
        if (replacement) {
          const changedAssignment = { ...assignment, teacherId: replacement.id }
          cleaned.push(changedAssignment)
          markChangedPair(slot, changedAssignment)
          changeLog.push({ slot, action: '講師差替', detail: `${replacement.name} に変更（元の講師が削除済）` })
        } else {
          changeLog.push({ slot, action: '講師削除', detail: `割当解除（講師が削除済・代替不可）` })
        }
        continue
      }

      // Check teacher still has availability for this slot
      if (!hasAvailability(data.availability, 'teacher', assignment.teacherId, slot)) {
        const teacherName = data.teachers.find((t) => t.id === assignment.teacherId)?.name ?? '?'
        const usedTeachers = new Set(cleaned.map((a) => a.teacherId))
        const replacement = data.teachers.find((t) => {
          if (usedTeachers.has(t.id)) return false
          if (!hasAvailability(data.availability, 'teacher', t.id, slot)) return false
          return t.subjects.includes(assignment.subject)
        })
        if (replacement) {
          const changedAssignment = { ...assignment, teacherId: replacement.id }
          cleaned.push(changedAssignment)
          markChangedPair(slot, changedAssignment)
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
          changeLog.push({ slot, action: '生徒解除', detail: `${studentName} を解除（予定変更のため）` })
        }
      }

      if (validStudentIds.length > 0) {
        const changedAssignment = { ...assignment, studentIds: validStudentIds }
        cleaned.push(changedAssignment)
        if (assignment.studentIds.length !== validStudentIds.length) {
          markChangedPair(slot, changedAssignment)
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
        const key = `${studentId}|${assignment.subject}`
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
        const requested = student?.subjectSlots[assignment.subject] ?? 0
        const key = `${studentId}|${assignment.subject}`
        const currentLoad = specialLoadMap.get(key) ?? 0

        if (currentLoad > requested) {
          specialLoadMap.set(key, currentLoad - 1)
          removedAny = true
          changeLog.push({ slot, action: '希望減で解除', detail: `${student?.name ?? studentId} (${assignment.subject})` })
          continue
        }
        remainingStudentIds.push(studentId)
      }

      if (removedAny) {
        assignment.studentIds = remainingStudentIds
        if (remainingStudentIds.length > 0) {
          markChangedPair(slot, assignment)
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
        if (!student.subjects.includes(assignment.subject)) return false
        const requested = student.subjectSlots[assignment.subject] ?? 0
        const allocated = countStudentSubjectLoad(result, student.id, assignment.subject)
        return allocated < requested
      })

      if (candidates.length > 0 && assignment.studentIds.length < 2) {
        const best = candidates.sort((a, b) => {
          const aRem = Object.values(a.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, a.id)
          const bRem = Object.values(b.subjectSlots).reduce((s, c) => s + c, 0) - countStudentLoad(result, b.id)
          return bRem - aRem
        })[0]
        assignment.studentIds = [...assignment.studentIds, best.id]
        markChangedPair(slot, assignment)
        changeLog.push({ slot, action: '生徒追加', detail: `${best.name} を追加` })
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
        return teacher.subjects.some((subject) => student.subjects.includes(subject))
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

        const commonSubjects = teacher.subjects.filter((subject) =>
          combo.every((student) => student.subjects.includes(subject)),
        )
        if (commonSubjects.length === 0) continue

        const viableSubjects = commonSubjects.filter((subject) =>
          combo.every((student) => {
            const requested = student.subjectSlots[subject] ?? 0
            const allocated = countStudentSubjectLoad(result, student.id, subject)
            return allocated < requested
          }),
        )
        if (viableSubjects.length === 0) continue

        const bestSubject = viableSubjects[0]

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

        const score = 100 +
          (isExistingDate ? 80 : -50) +  // Very strong preference for reusing existing dates
          teacherConsecutiveBonus +  // Teacher consecutive slot bonus
          firstHalfBonus +  // First-half bias (max +25)
          regularPairBonus +  // Regular lesson pair preference
          pairConsecutiveBonus +  // Same teacher+student consecutive on same day
          (combo.length === 2 ? 30 : 0) +  // 2-person pair bonus
          studentScore -
          teacherLoad * 2

        if (!bestPlan || score > bestPlan.score) {
          bestPlan = { score, assignment: { teacherId: teacher.id, studentIds: combo.map((s) => s.id), subject: bestSubject } }
        }
      }

      if (bestPlan) {
        slotAssignments.push(bestPlan.assignment)
        markChangedPair(slot, bestPlan.assignment)
        usedTeacherIdsInSlot.add(teacher.id)
        for (const sid of bestPlan.assignment.studentIds) usedStudentIdsInSlot.add(sid)
      }
    }

    if (slotAssignments.length > existingAssignments.length) {
      result[slot] = slotAssignments
      // Only log newly added assignments
      for (const a of slotAssignments.slice(existingAssignments.length)) {
        markAddedPair(slot, a)
        const tName = data.teachers.find((t) => t.id === a.teacherId)?.name ?? '?'
        const sNames = a.studentIds.map((sid) => data.students.find((s) => s.id === sid)?.name ?? '?').join(', ')
        changeLog.push({ slot, action: '新規割当', detail: `${tName} × ${sNames} (${a.subject})` })
      }
    }
  }

  const changedPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(changedPairSigSetBySlot)) {
    if (signatureSet.size > 0) changedPairSignatures[slot] = [...signatureSet]
  }

  const addedPairSignatures: Record<string, string[]> = {}
  for (const [slot, signatureSet] of Object.entries(addedPairSigSetBySlot)) {
    if (signatureSet.size > 0) addedPairSignatures[slot] = [...signatureSet]
  }

  return { assignments: result, changeLog, changedPairSignatures, addedPairSignatures }
}

const emptyMasterData = (): MasterData => ({
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
  const [newTerm, setNewTerm] = useState<'spring' | 'summer' | 'winter'>('summer')
  const [newSessionId, setNewSessionId] = useState('')
  const [newSessionName, setNewSessionName] = useState('')
  const [newStartDate, setNewStartDate] = useState('')
  const [newEndDate, setNewEndDate] = useState('')
  const [newSubmissionStart, setNewSubmissionStart] = useState('')
  const [newSubmissionEnd, setNewSubmissionEnd] = useState('')
  const [newDeskCount, setNewDeskCount] = useState(0)
  const [newHolidays, setNewHolidays] = useState<string[]>([])

  // Master data form state
  const [teacherName, setTeacherName] = useState('')
  const [teacherSubjects, setTeacherSubjects] = useState<string[]>([])
  const [teacherMemo, setTeacherMemo] = useState('')
  const [studentName, setStudentName] = useState('')
  const [studentGrade, setStudentGrade] = useState('')
  const [constraintTeacherId, setConstraintTeacherId] = useState('')
  const [constraintStudentId, setConstraintStudentId] = useState('')
  const [constraintType, setConstraintType] = useState<ConstraintType>('incompatible')
  const [gradeConstraintTeacherId, setGradeConstraintTeacherId] = useState('')
  const [gradeConstraintGrade, setGradeConstraintGrade] = useState('')
  const [gradeConstraintType, setGradeConstraintType] = useState<ConstraintType>('incompatible')
  const [regularTeacherId, setRegularTeacherId] = useState('')
  const [regularStudent1Id, setRegularStudent1Id] = useState('')
  const [regularStudent2Id, setRegularStudent2Id] = useState('')
  const [regularSubject, setRegularSubject] = useState('')
  const [regularDayOfWeek, setRegularDayOfWeek] = useState('')
  const [regularSlotNumber, setRegularSlotNumber] = useState('')
  const fileInputRef = useRef<HTMLInputElement>(null)

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
    const label = newTerm === 'spring' ? '春期講習' : newTerm === 'summer' ? '夏期講習' : '冬期講習'
    const idTerm = newTerm === 'spring' ? 'spring' : newTerm === 'summer' ? 'summer' : 'winter'
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

  const addTeacher = async (): Promise<void> => {
    if (!teacherName.trim()) return
    const teacher: Teacher = { id: createId(), name: teacherName.trim(), subjects: teacherSubjects, memo: teacherMemo.trim() }
    await updateMaster((c) => ({ ...c, teachers: [...c.teachers, teacher] }))
    setTeacherName(''); setTeacherSubjects([]); setTeacherMemo('')
  }

  const addStudent = async (): Promise<void> => {
    if (!studentName.trim()) return
    const student: Student = {
      id: createId(), name: studentName.trim(), grade: studentGrade.trim(),
      subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo: '', submittedAt: 0,
    }
    await updateMaster((c) => ({ ...c, students: [...c.students, student] }))
    setStudentName(''); setStudentGrade('')
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
    const nc: GradeConstraint = { id: createId(), teacherId: gradeConstraintTeacherId, grade: gradeConstraintGrade, type: gradeConstraintType }
    await updateMaster((c) => {
      const filtered = (c.gradeConstraints ?? []).filter((i) => !(i.teacherId === gradeConstraintTeacherId && i.grade === gradeConstraintGrade))
      return { ...c, gradeConstraints: [...filtered, nc] }
    })
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
      ['田中講師', '数,英', '数学メイン'],
      ['佐藤講師', '英,数', '英語メイン'],
    ]
    const sampleStudents = [
      ['青木 太郎', '中3'],
      ['伊藤 花', '中2'],
      ['上田 陽介', '高1'],
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
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '担当科目(カンマ区切り: ' + FIXED_SUBJECTS.join(',') + ')', 'メモ'], ...sampleTeachers]), '講師')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '学年'], ...sampleStudents]), '生徒')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒名', '種別(不可)'], ...sampleConstraints]), '制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '学年', '種別(不可)'], ...sampleGradeConstraints]), '学年制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒1名', '生徒2名(任意)', '科目', '曜日(月/火/水/木/金/土/日)', '時限番号'], ...sampleRegularLessons]), '通常授業')
    XLSX.writeFile(wb, 'テンプレート.xlsx')
  }

  const exportData = (): void => {
    if (!masterData) return
    const md = masterData
    const teacherRows = md.teachers.map((t) => [t.name, t.subjects.join(', '), t.memo])
    const studentRows = md.students.map((s) => [s.name, s.grade, s.memo])
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
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '担当科目', 'メモ'], ...teacherRows]), '講師')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '学年', 'メモ'], ...studentRows]), '生徒')
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
        if (md.teachers.some((t) => t.name === name)) continue
        importedTeachers.push({ id: createId(), name, subjects, memo })
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
        if (md.students.some((s) => s.name === name)) continue
        importedStudents.push({ id: createId(), name, grade, subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo: '', submittedAt: 0 })
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
    if (!masterData || (masterData.teachers.length === 0 && masterData.students.length === 0)) {
      alert('管理データ（講師・生徒）が未登録です。先に管理データを登録してください。')
      return
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
    seed.teachers = masterData.teachers
    seed.students = masterData.students.map((s) => ({
      ...s, subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], submittedAt: 0,
    }))
    seed.constraints = masterData.constraints
    seed.gradeConstraints = masterData.gradeConstraints
    seed.regularLessons = masterData.regularLessons
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
              {(!masterData || (masterData.teachers.length === 0 && masterData.students.length === 0)) ? (
                <p style={{ color: '#dc2626', fontWeight: 600 }}>⚠ 管理データ（講師・生徒）が未登録のため、特別講習を追加できません。先に下部の管理データを登録してください。</p>
              ) : (
                <>
                  <p className="muted">作成時にマスターデータ（講師・生徒・制約・通常授業）が自動コピーされます。</p>
                  <div className="row">
                    <input value={newYear} onChange={(e) => setNewYear(e.target.value)} placeholder="西暦" style={{ width: 80 }} />
                    <select value={newTerm} onChange={(e) => setNewTerm(e.target.value as typeof newTerm)}>
                      <option value="spring">春期講習</option>
                      <option value="summer">夏期講習</option>
                      <option value="winter">冬期講習</option>
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
                  <h3>講師登録</h3>
                  <div className="row">
                    <input value={teacherName} onChange={(e) => setTeacherName(e.target.value)} placeholder="講師名" />
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
                    <thead><tr><th>名前</th><th>科目</th><th>メモ</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.teachers.map((t) => (
                        <tr key={t.id}><td>{t.name}</td><td>{t.subjects.join(', ')}</td><td>{t.memo}</td>
                          <td><button className="btn secondary" type="button" onClick={() => void removeTeacher(t.id)}>削除</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>生徒登録</h3>
                  <div className="row">
                    <input value={studentName} onChange={(e) => setStudentName(e.target.value)} placeholder="生徒名" />
                    <select value={studentGrade} onChange={(e) => setStudentGrade(e.target.value)}>
                      <option value="">学年を選択</option>
                      {GRADE_OPTIONS.map((g) => (<option key={g} value={g}>{g}</option>))}
                    </select>
                    <button className="btn" type="button" onClick={() => void addStudent()}>追加</button>
                  </div>
                  <table className="table">
                    <thead><tr><th>名前</th><th>学年</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.students.map((s) => (
                        <tr key={s.id}><td>{s.name}</td><td>{s.grade}</td>
                          <td><button className="btn secondary" type="button" onClick={() => void removeStudent(s.id)}>削除</button></td>
                        </tr>
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
                  <div className="row">
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
                    <button className="btn" type="button" onClick={() => void upsertGradeConstraint()}>保存</button>
                  </div>
                  <table className="table">
                    <thead><tr><th>講師</th><th>学年</th><th>種別</th><th>操作</th></tr></thead>
                    <tbody>
                      {(masterData.gradeConstraints ?? []).map((gc) => (
                        <tr key={gc.id}>
                          <td>{masterData.teachers.find((t) => t.id === gc.teacherId)?.name ?? '-'}</td>
                          <td>{gc.grade}</td>
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

const AdminPage = () => {
  const { sessionId = 'main' } = useParams()
  const navigate = useNavigate()
  const location = useLocation()
  const skipAuth = (location.state as { skipAuth?: boolean } | null)?.skipAuth === true
  const { data, setData, loading, error: sessionError } = useSessionData(sessionId)
  const [authorized, setAuthorized] = useState(import.meta.env.DEV || skipAuth)
  const [recentlyUpdated, setRecentlyUpdated] = useState<Set<string>>(new Set())
  const [dragInfo, setDragInfo] = useState<{ sourceSlot: string; sourceIdx: number; teacherId: string; studentIds: string[] } | null>(null)
  const prevSnapshotRef = useRef<{ availability: Record<string, string[]>; studentSubmittedAt: Record<string, number> } | null>(null)

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
    // Check teachers: availability change
    for (const teacher of data.teachers) {
      const key = `teacher:${teacher.id}`
      const prevSlots = (prev.availability[key] ?? []).slice().sort().join(',')
      const currSlots = (data.availability[key] ?? []).slice().sort().join(',')
      if (prevSlots !== currSlots) updatedIds.push(teacher.id)
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

  const openInputPage = async (personType: PersonType, personId: string): Promise<void> => {
    navigate(`/availability/${sessionId}/${personType}/${personId}`, { state: { fromAdminInput: true } })
  }

  useEffect(() => {
    setAuthorized(import.meta.env.DEV || skipAuth)
  }, [sessionId, skipAuth])

  const slotKeys = useMemo(() => (data ? buildSlotKeys(data.settings) : []), [data])

  const persist = async (next: SessionData): Promise<void> => {
    setData(next)
    await saveSession(sessionId, next)
  }

  const update = async (updater: (current: SessionData) => SessionData): Promise<void> => {
    if (!data) return
    await persist(updater(data))
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
        teachers: master.teachers,
        students: mergedStudents,
        constraints: master.constraints,
        gradeConstraints: master.gradeConstraints,
        regularLessons: master.regularLessons,
      }
      // Only save if something actually changed
      const changed =
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

  // Auto-fill regular lessons when date range or regular lessons change
  const regularFillSigRef = useRef('')
  useEffect(() => {
    if (!data || !authorized || slotKeys.length === 0) return
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

    if (changed) {
      void persist({ ...data, assignments: nextAssignments })
    }
  }, [slotKeys, data, authorized]) // eslint-disable-line react-hooks/exhaustive-deps

  const applyAutoAssign = async (): Promise<void> => {
    if (!data) return
    const hadNonRegularBefore = Object.values(data.assignments).some((slotAssignments) =>
      slotAssignments.some((assignment) => hasMeaningfulManualAssignment(assignment)),
    )
    const { assignments: nextAssignments, changeLog, changedPairSignatures, addedPairSignatures } = buildIncrementalAutoAssignments(data, slotKeys)

    const mergedHighlights: Record<string, string[]> = {}
    if (hadNonRegularBefore) {
      const slotSet = new Set([...Object.keys(addedPairSignatures), ...Object.keys(changedPairSignatures)])
      for (const slot of slotSet) {
        const signatureSet = new Set<string>([
          ...(addedPairSignatures[slot] ?? []),
          ...(changedPairSignatures[slot] ?? []),
        ])
        if (signatureSet.size > 0) mergedHighlights[slot] = [...signatureSet]
      }
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

    await update((current) => ({
      ...current,
      assignments: nextAssignments,
      autoAssignHighlights: hadNonRegularBefore ? mergedHighlights : {},
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
          .map((item) => `${slotLabel(item.slot)} ${item.detail}`)
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
    await update((current) => ({
      ...current,
      assignments: {},
      autoAssignHighlights: {},
    }))
  }

  /** Export schedule to Excel: weekly sheets (Mon–Sun) */
  const exportScheduleExcel = (): void => {
    if (!data) return
    const { startDate, endDate, slotsPerDay, holidays } = data.settings
    if (!startDate || !endDate) { alert('開始日・終了日を設定してください。'); return }

    const holidaySet = new Set(holidays)
    const dayNames = ['日', '月', '火', '水', '木', '金', '土']

    // Build all dates in range
    const allDates: string[] = []
    const start = new Date(`${startDate}T00:00:00`)
    const end = new Date(`${endDate}T00:00:00`)
    for (let cursor = new Date(start); cursor <= end; cursor = new Date(cursor.getTime() + 86400000)) {
      const y = cursor.getFullYear()
      const m = String(cursor.getMonth() + 1).padStart(2, '0')
      const d = String(cursor.getDate()).padStart(2, '0')
      allDates.push(`${y}-${m}-${d}`)
    }

    if (allDates.length === 0) { alert('期間に日付がありません。'); return }

    // Group dates into weeks (Mon–Sun)
    const weeks: string[][] = []
    let currentWeek: string[] = []
    for (const date of allDates) {
      const dow = getIsoDayOfWeek(date)
      if (dow === 1 && currentWeek.length > 0) {
        weeks.push(currentWeek)
        currentWeek = []
      }
      currentWeek.push(date)
    }
    if (currentWeek.length > 0) weeks.push(currentWeek)

    // Build cell text for a slot assignment
    const buildCellText = (slotKey: string): string => {
      const slotAssignments = data.assignments[slotKey] ?? []
      if (slotAssignments.length === 0) return ''
      return slotAssignments.map((a) => {
        const tName = data.teachers.find((t) => t.id === a.teacherId)?.name ?? ''
        const sNames = a.studentIds.map((sid) => data.students.find((st) => st.id === sid)?.name ?? '').join(', ')
        const regular = a.isRegular ? '[通常] ' : ''
        return `${regular}${tName} / ${sNames} (${a.subject})`
      }).join(' | ')
    }

    try {
      const wb = XLSX.utils.book_new()

      for (let wi = 0; wi < weeks.length; wi++) {
        const weekDates = weeks[wi]
        const firstDate = weekDates[0]
        const lastDate = weekDates[weekDates.length - 1]
        const [, fm, fd] = firstDate.split('-')
        const [, lm, ld] = lastDate.split('-')

        // Pad to full Mon–Sun week
        const fullWeek: (string | null)[] = []
        const firstDow = getIsoDayOfWeek(firstDate)
        const startPad = firstDow === 0 ? 6 : firstDow - 1
        for (let p = 0; p < startPad; p++) fullWeek.push(null)
        for (const d of weekDates) fullWeek.push(d)
        while (fullWeek.length < 7) fullWeek.push(null)

        // Header row
        const dowOrder = [1, 2, 3, 4, 5, 6, 0]
        const header: string[] = ['']
        for (let i = 0; i < 7; i++) {
          const date = fullWeek[i]
          if (!date) {
            header.push(`${dayNames[dowOrder[i]]}`)
          } else {
            const [, mm, dd] = date.split('-')
            header.push(`${Number(mm)}/${Number(dd)}(${dayNames[dowOrder[i]]})`)
          }
        }

        // Data rows
        const rows: string[][] = []
        for (let s = 1; s <= slotsPerDay; s++) {
          const row: string[] = [`${s}限`]
          for (let i = 0; i < 7; i++) {
            const date = fullWeek[i]
            if (!date) {
              row.push('')
            } else if (holidaySet.has(date)) {
              row.push('休')
            } else {
              row.push(buildCellText(`${date}_${s}`))
            }
          }
          rows.push(row)
        }

        const aoa = [header, ...rows]
        const ws = XLSX.utils.aoa_to_sheet(aoa)

        // Column widths only (community XLSX does not support cell styling)
        ws['!cols'] = [{ wch: 5 }, ...Array(7).fill({ wch: 36 })]

        const sheetName = `${Number(fm)}月${Number(fd)}日-${Number(lm)}月${Number(ld)}日`
        XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31))
      }

      XLSX.writeFile(wb, `コマ割り_${data.settings.name}.xlsx`)
    } catch (err) {
      console.error('Excel export error:', err)
      alert('Excel出力に失敗しました: ' + String(err))
    }
  }

  const setSlotTeacher = async (slot: string, idx: number, teacherId: string): Promise<void> => {
    await update((current) => {
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
      const currentTeacher = current.teachers.find((item) => item.id === teacherId)
      const nextSubject =
        prev?.subject && currentTeacher?.subjects.includes(prev.subject)
          ? prev.subject
          : (currentTeacher?.subjects[0] ?? '')

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
    await update((current) => {
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
    await update((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      const assignment = slotAssignments[idx]
      if (!assignment) {
        return current
      }

      const prevIds = [...assignment.studentIds]
      if (studentId === '') {
        prevIds.splice(position, 1)
      } else {
        prevIds[position] = studentId
      }
      const studentIds = prevIds.filter(Boolean)

      const teacher = current.teachers.find((item) => item.id === assignment.teacherId)
      const commonSubjects = (teacher?.subjects ?? []).filter((subject) =>
        studentIds.every((id) => current.students.find((student) => student.id === id)?.subjects.includes(subject)),
      )

      slotAssignments[idx] = {
        ...assignment,
        studentIds,
        subject: commonSubjects.includes(assignment.subject)
          ? assignment.subject
          : (commonSubjects[0] ?? assignment.subject),
      }

      return {
        ...current,
        assignments: { ...current.assignments, [slot]: slotAssignments },
      }
    })
  }

  const setSlotSubject = async (slot: string, idx: number, subject: string): Promise<void> => {
    await update((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      const assignment = slotAssignments[idx]
      if (!assignment) {
        return current
      }
      slotAssignments[idx] = { ...assignment, subject }
      return {
        ...current,
        assignments: { ...current.assignments, [slot]: slotAssignments },
      }
    })
  }

  /** Move an assignment from one slot to another (drag-and-drop) */
  const moveAssignment = async (sourceSlot: string, sourceIdx: number, targetSlot: string): Promise<void> => {
    await update((current) => {
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
      if (moved.teacherId && !hasAvailability(current.availability, 'teacher', moved.teacherId, targetSlot)) return current

      // Check all assigned students are available in target slot
      const movedStudents = current.students.filter((s) => moved.studentIds.includes(s.id))
      if (movedStudents.some((student) => !isStudentAvailable(student, targetSlot))) return current

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
            <h3>講師一覧</h3>
            <table className="table">
              <thead><tr><th>名前</th><th>科目</th><th>提出データ</th><th>代行入力</th><th>共有</th></tr></thead>
              <tbody>
                {data.teachers.map((teacher) => {
                  const teacherSubmittedAt = (data.teacherSubmittedAt ?? {})[teacher.id] ?? 0
                  return (
                  <tr key={teacher.id}>
                    <td>
                      {teacher.name}
                      {recentlyUpdated.has(teacher.id) && (
                        <span className="badge ok" style={{ marginLeft: '8px', fontSize: '11px', animation: 'fadeIn 0.3s' }}>✓ 更新済</span>
                      )}
                    </td>
                    <td>{teacher.subjects.join(', ')}</td>
                    <td>
                      {teacherSubmittedAt ? (
                        <span style={{ fontSize: '0.85em', color: '#16a34a' }}>
                          {new Date(teacherSubmittedAt).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' })}
                          {' '}提出済
                        </span>
                      ) : (
                        <span style={{ fontSize: '0.85em', color: '#dc2626', fontWeight: 600 }}>未提出</span>
                      )}
                    </td>
                    <td><button className="btn secondary" type="button" onClick={() => void openInputPage('teacher', teacher.id)}>入力ページ</button></td>
                    <td><button className="btn secondary" type="button" onClick={() => void copyInputUrl('teacher', teacher.id)}>URLコピー</button></td>
                  </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <h3>生徒一覧</h3>
            <p className="muted">希望コマ数・不可日は生徒本人が希望URLから入力します。</p>
            <table className="table">
              <thead><tr><th>名前</th><th>学年</th><th>提出データ</th><th>代行入力</th><th>共有</th></tr></thead>
              <tbody>
                {data.students.map((student) => (
                  <tr key={student.id}>
                    <td>
                      {student.name}
                      {recentlyUpdated.has(student.id) && (
                        <span className="badge ok" style={{ marginLeft: '8px', fontSize: '11px', animation: 'fadeIn 0.3s' }}>✓ 更新済</span>
                      )}
                    </td>
                    <td>{student.grade}</td>
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
                    <td><button className="btn secondary" type="button" onClick={() => void copyInputUrl('student', student.id)}>URLコピー</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <div className="row">
              <h3>コマ割り</h3>
              {(() => {
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
                  .map((item) => `${slotLabel(item.slot)}: ${item.detail}`)
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
                        講師不足: {teacherShortages.length}件
                      </span>
                    )}
                  </>
                )
              })()}
              <button className="btn secondary" type="button" onClick={() => void applyAutoAssign()}>
                自動提案
              </button>
              <button className="btn secondary" type="button" onClick={() => void resetAssignments()}>
                コマ割りリセット
              </button>
              <button className="btn" type="button" onClick={exportScheduleExcel}>
                Excel出力
              </button>
            </div>
            <p className="muted">通常授業は日付確定時に自動配置。特別講習は自動提案で割当。講師1人 + 生徒1〜2人。</p>
            <p className="muted" style={{ fontSize: '12px' }}>★=通常授業　⚠=制約不可　ペアはドラッグで別コマへ移動可</p>
            <div className="grid-slots">
              {slotKeys.map((slot) => {
                const slotAssignments = data.assignments[slot] ?? []
                const usedTeacherIds = new Set(slotAssignments.map((a) => a.teacherId).filter(Boolean))

                // D&D: compute validity of this slot as a drop target
                const deskCount = data.settings.deskCount ?? 0
                const isDragActive = dragInfo !== null
                const isSameSlot = isDragActive && dragInfo.sourceSlot === slot
                const isDeskFull = isDragActive && deskCount > 0 && slotAssignments.length >= deskCount
                const isTeacherConflict = isDragActive && dragInfo.teacherId && usedTeacherIds.has(dragInfo.teacherId)
                const draggedStudents = isDragActive ? data.students.filter((s) => dragInfo.studentIds.includes(s.id)) : []
                const hasUnavailableStudent = isDragActive && draggedStudents.some((student) => !isStudentAvailable(student, slot))
                const hasStudentConflict = isDragActive && dragInfo.studentIds.some((sid) => slotAssignments.some((a) => a.studentIds.includes(sid)))
                const hasTeacherUnavailable = isDragActive && dragInfo.teacherId ? !hasAvailability(data.availability, 'teacher', dragInfo.teacherId, slot) : false
                const isDropValid = isDragActive && !isSameSlot && !isDeskFull && !isTeacherConflict && !hasUnavailableStudent && !hasStudentConflict && !hasTeacherUnavailable
                const slotDragClass = isDragActive ? (isSameSlot ? '' : isDropValid ? ' drag-valid' : ' drag-invalid') : ''

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
                        const { sourceSlot, sourceIdx } = JSON.parse(raw) as { sourceSlot: string; sourceIdx: number }
                        if (sourceSlot === slot) {
                          setDragInfo(null)
                          return
                        }
                        void moveAssignment(sourceSlot, sourceIdx, slot)
                      } catch { /* ignore */ }
                      setDragInfo(null)
                    }}
                  >
                    <div className="slot-title">
                      <div>
                        {slotLabel(slot)}
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
                        const selectedTeacher = data.teachers.find((t) => t.id === assignment.teacherId)
                        const selectedStudents = data.students.filter((s) =>
                          assignment.studentIds.includes(s.id),
                        )
                        const subjectOptions = selectedTeacher
                          ? selectedTeacher.subjects.filter((subject) =>
                              selectedStudents.length === 0
                                ? true
                                : selectedStudents.every((s) => s.subjects.includes(subject)),
                            )
                          : []

                        const isIncompatiblePair = assignment.teacherId && selectedStudents.some((s) => {
                          const pt = constraintFor(data.constraints, assignment.teacherId, s.id)
                          const gt = gradeConstraintFor(data.gradeConstraints ?? [], assignment.teacherId, s.grade)
                          return pt === 'incompatible' || gt === 'incompatible'
                        })
                        const isAutoDiff = (data.autoAssignHighlights?.[slot] ?? []).includes(assignmentSignature(assignment))

                        return (
                          <div
                            key={idx}
                            className={`assignment-block${assignment.isRegular ? ' assignment-block-regular' : ''}${isIncompatiblePair ? ' assignment-block-incompatible' : ''}${isAutoDiff ? ' assignment-block-auto-updated' : ''}`}
                            draggable={!assignment.isRegular}
                            onDragStart={(e) => {
                              const payload = JSON.stringify({ sourceSlot: slot, sourceIdx: idx })
                              e.dataTransfer.setData('text/plain', payload)
                              e.dataTransfer.effectAllowed = 'move'
                              setDragInfo({ sourceSlot: slot, sourceIdx: idx, teacherId: assignment.teacherId, studentIds: [...assignment.studentIds] })
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
                            {isAutoDiff && <span className="badge auto-diff-badge" title="自動提案で差分あり">NEW</span>}
                            <select
                              value={assignment.teacherId}
                              onChange={(e) => void setSlotTeacher(slot, idx, e.target.value)}
                              disabled={assignment.isRegular}
                            >
                              <option value="">講師を選択</option>
                              {data.teachers
                                .filter((teacher) => {
                                  // Always show currently assigned teacher
                                  if (teacher.id === assignment.teacherId) return true
                                  // Only show teachers who have availability for this slot
                                  return hasAvailability(data.availability, 'teacher', teacher.id, slot)
                                })
                                .map((teacher) => {
                                const usedElsewhere = usedTeacherIds.has(teacher.id) && teacher.id !== assignment.teacherId
                                return (
                                  <option key={teacher.id} value={teacher.id} disabled={usedElsewhere}>
                                    {teacher.name}{usedElsewhere ? ' (割当済)' : ''}
                                  </option>
                                )
                              })}
                            </select>

                            {assignment.teacherId && (
                              <>
                                <select
                                  value={assignment.subject}
                                  onChange={(e) => void setSlotSubject(slot, idx, e.target.value)}
                                  disabled={assignment.isRegular}
                                >
                                  {subjectOptions.map((subject) => (
                                    <option key={subject} value={subject}>
                                      {subject}
                                    </option>
                                  ))}
                                </select>

                                {[0, 1].map((pos) => {
                                  const otherStudentId = assignment.studentIds[pos === 0 ? 1 : 0] ?? ''
                                  return (
                                    <div key={pos} className="student-select-row">
                                    <select
                                      value={assignment.studentIds[pos] ?? ''}
                                      disabled={assignment.isRegular}
                                      onChange={(e) => {
                                        const selectedId = e.target.value
                                        if (selectedId) {
                                          const student = data.students.find((s) => s.id === selectedId)
                                          if (student) {
                                            const pairTag = constraintFor(data.constraints, assignment.teacherId, student.id)
                                            const gradeTag = gradeConstraintFor(data.gradeConstraints ?? [], assignment.teacherId, student.grade)
                                            const isIncompatible = pairTag === 'incompatible' || gradeTag === 'incompatible'
                                            if (isIncompatible) {
                                              const reasons: string[] = []
                                              if (pairTag === 'incompatible') reasons.push('ペア制約で不可')
                                              if (gradeTag === 'incompatible') reasons.push(`学年制約(${student.grade})で不可`)
                                              const ok = window.confirm(
                                                `⚠️ ${student.name} は制約ルールにより割当不可です。\n理由: ${reasons.join(', ')}\n\nそれでも割り当てますか？`,
                                              )
                                              if (!ok) {
                                                e.target.value = assignment.studentIds[pos] ?? ''
                                                return
                                              }
                                            }
                                          }
                                        }
                                        void setSlotStudent(slot, idx, pos, selectedId)
                                      }}
                                    >
                                      <option value="">{`生徒${pos + 1}を選択`}</option>
                                      {data.students
                                        .filter((student) => {
                                          // Always show currently assigned student
                                          if (student.id === assignment.studentIds[pos]) return true
                                          // Unsubmitted students are unavailable
                                          if (!student.submittedAt) return false
                                          // Filter out students unavailable for this specific slot
                                          if (!isStudentAvailable(student, slot)) return false
                                          // Subject compatibility is required
                                          if (!assignment.subject || !student.subjects.includes(assignment.subject)) return false
                                          // Remaining slots for the current subject must be > 0
                                          const desired = student.subjectSlots[assignment.subject] ?? 0
                                          const assigned = countStudentSubjectLoad(data.assignments, student.id, assignment.subject)
                                          const currentlySelected = assignment.studentIds[pos] === student.id && assignment.subject
                                          const adjustedAssigned = currentlySelected ? Math.max(0, assigned - 1) : assigned
                                          const remaining = desired - adjustedAssigned
                                          return remaining > 0
                                        })
                                        .map((student) => {
                                        const pairTag = constraintFor(data.constraints, assignment.teacherId, student.id)
                                        const gradeTag = gradeConstraintFor(data.gradeConstraints ?? [], assignment.teacherId, student.grade)
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
                                    </div>
                                  )
                                })}
                              </>
                            )}

                          </div>
                        )
                      })}
                      {(() => {
                        const idleTeachers = data.teachers.filter(
                          (t) =>
                            hasAvailability(data.availability, 'teacher', t.id, slot) &&
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
const TeacherInputPage = ({
  sessionId,
  data,
  teacher,
  returnToAdminOnComplete,
}: {
  sessionId: string
  data: SessionData
  teacher: Teacher
  returnToAdminOnComplete: boolean
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
    const key = personKey('teacher', teacher.id)
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
    const key = personKey('teacher', teacher.id)
    // Ensure regular lesson slots are always included
    const merged = new Set(localAvailability)
    for (const rk of regularSlotKeys) merged.add(rk)
    const availabilityArray = Array.from(merged)

    // Determine if this is initial or update submission
    const isUpdate = !!(data.teacherSubmittedAt?.[teacher.id])
    const logEntry: SubmissionLogEntry = {
      personId: teacher.id,
      personType: 'teacher',
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
    navigate(`/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
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
                  {i + 1}限
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
    navigate(`/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
  }

  return (
    <div className="availability-container">
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
                    {i + 1}限
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

const AvailabilityPage = () => {
  const location = useLocation()
  const { sessionId = 'main', personType: rawPersonType = 'teacher', personId: rawPersonId = '' } = useParams()
  const personType = (rawPersonType === 'student' ? 'student' : 'teacher') as PersonType
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
      return <StudentInputPage sessionId={sessionId} data={data} student={currentPerson as Student} returnToAdminOnComplete={returnToAdminOnComplete} />
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
