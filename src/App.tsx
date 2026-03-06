import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { Link, Route, Routes, useLocation, useNavigate, useParams } from 'react-router-dom'
import XLSX from 'xlsx-js-style'
import './App.css'
import { cleanupOldBackups, createBackup, createClassroom, deleteBackup, deleteClassroom, deleteSession, findClassroomForSession, initAuth, listBackups, listSessionItems, loadBackup, loadMasterData, loadSession, restoreBackup, saveAndVerify, saveMasterData, saveSession, watchClassrooms, watchMasterData, watchSession, watchSessionsList, type BackupMeta, type ClassroomInfo } from './firebase'
import type {
  ActualResult,
  Assignment,
  ConstraintCardType,
  ConstraintType,
  GroupLesson,
  Manager,
  MasterData,
  PairConstraint,
  PairConstraintPersonType,
  PersonType,
  RegularLesson,
  SessionData,
  Student,
  SubmissionLogEntry,
  Teacher,
} from './types'
import { buildSlotKeys, formatShortDate, mendanTimeLabel, personKey, slotLabel } from './utils/schedule'
import { BASE_SUBJECTS, TEACHER_SUBJECTS, canTeachSubject, teachableBaseSubjects, teacherHasSubject, getSubjectBase } from './utils/subjects'
import { downloadEmailReceiptPdf, downloadSubmissionReceiptPdf, exportSchedulePdf } from './utils/pdf'
import { constraintFor, hasAvailability, isStudentAvailable, isParentAvailableForMendan } from './utils/constraints'
import { getSlotNumber, getIsoDayOfWeek, getSlotDayOfWeek, buildEffectiveAssignments, getStudentSubject, countStudentSubjectLoad, collectTeacherShortages, assignmentSignature, hasMeaningfulManualAssignment, findRegularLessonsForSlot, getDatesInRange } from './utils/assignments'
import { buildIncrementalAutoAssignments, buildMendanAutoAssignments } from './utils/autoAssign'
import { ALL_CONSTRAINT_CARDS, CONSTRAINT_CARD_LABELS, CONSTRAINT_CARD_DESCRIPTIONS, CONSTRAINT_CARD_CONFLICT_GROUP, DAILY_LIMIT_CONFLICT_GROUP, evaluateConstraintCards, getDefaultConstraintCards, summarizeConstraintCards, validateConstraintCards } from './utils/slotConstraints'

const APP_VERSION = '1.3.21'

type ForceAssignAction = {
  type: 'force-assign'
  slot: string
  teacherId: string
  studentId: string
  subject: string
  makeupInfo?: { dayOfWeek: number; slotNumber: number; date?: string }
}

type StatusProposal = {
  label: string
  action?: ForceAssignAction
}

type StatusDetail = {
  label: string
  causes: string[]
  proposals: StatusProposal[]
}

type StatusSection = {
  key: 'under' | 'makeup' | 'over' | 'shortage' | 'overRemoved'
  title: string
  items: StatusDetail[]
}

type StatusReport = {
  title: string
  summary: string
  sections: StatusSection[]
}

type PendingMakeupDemand = {
  studentId: string
  teacherId: string
  subject: string
  absentDate?: string
  makeupInfo: { dayOfWeek: number; slotNumber: number; date?: string }
}

type PlacementAnalysis = {
  force: StatusProposal[]
  teacher: StatusProposal[]
  student: StatusProposal[]
  cards: StatusProposal[]
  blockers: string[]
}

const toStatusProposal = (label: string, action?: ForceAssignAction): StatusProposal => ({ label, ...(action ? { action } : {}) })

const dedupeStatusProposals = (proposals: StatusProposal[]): StatusProposal[] => {
  const seen = new Set<string>()
  const unique: StatusProposal[] = []
  for (const proposal of proposals) {
    const actionKey = proposal.action
      ? `${proposal.action.type}|${proposal.action.slot}|${proposal.action.teacherId}|${proposal.action.studentId}|${proposal.action.subject}|${proposal.action.makeupInfo?.dayOfWeek ?? ''}|${proposal.action.makeupInfo?.slotNumber ?? ''}|${proposal.action.makeupInfo?.date ?? ''}`
      : ''
    const key = `${proposal.label}|${actionKey}`
    if (seen.has(key)) continue
    seen.add(key)
    unique.push(proposal)
  }
  return unique
}

const GRADE_OPTIONS = ['小1', '小2', '小3', '小4', '小5', '小6', '中1', '中2', '中3', '高1', '高2', '高3']

const FIXED_SUBJECTS = ['英', '数', '国', '理', '社', 'IT', '算'] as readonly string[]
/** Leveled teacher subjects (小英, 中英, 高英, ...). */
const ALL_TEACHER_SUBJECTS = TEACHER_SUBJECTS

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
  subjects: [...FIXED_SUBJECTS],
  managers: [],
  teachers: [],
  students: [],
  constraints: [],
  gradeConstraints: [],
  availability: {},
  assignments: {},
  regularLessons: [],
  groupLessons: [],
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

  const subjects = [...FIXED_SUBJECTS]

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
    { id: 'c001', personAId: 't001', personBId: 's002', personAType: 'teacher', personBType: 'student', type: 'incompatible' },
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
    groupLessons: [],
  }
}

const useSessionData = (classroomId: string, sessionId: string) => {
  const [data, setData] = useState<SessionData | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    setLoading(true)
    setError(null)
    const unsub = watchSession(
      classroomId,
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
  }, [classroomId, sessionId])

  return { data, setData, loading, error }
}

const ADMIN_PASSWORD_STORAGE_KEY = 'lst_admin_password_v1'
const readSavedAdminPassword = (): string => localStorage.getItem(ADMIN_PASSWORD_STORAGE_KEY) ?? 'admin1234'
const saveAdminPassword = (password: string): void => {
  localStorage.setItem(ADMIN_PASSWORD_STORAGE_KEY, password)
}


const emptyMasterData = (): MasterData => ({
  managers: [],
  teachers: [],
  students: [],
  constraints: [],
  gradeConstraints: [],
  regularLessons: [],
  groupLessons: [],
})

/** Inline calendar for picking multiple holiday dates. */
const HolidayCalendar = ({ selected, onChange }: { selected: string[]; onChange: (dates: string[]) => void }) => {
  const today = new Date()
  const [year, setYear] = useState(today.getFullYear())
  const [month, setMonth] = useState(today.getMonth()) // 0-based

  const daysInMonth = new Date(year, month + 1, 0).getDate()
  const firstDow = new Date(year, month, 1).getDay() // 0=Sun
  const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`

  const pad2 = (n: number) => String(n).padStart(2, '0')
  const dateStr = (d: number) => `${year}-${pad2(month + 1)}-${pad2(d)}`

  const toggle = (d: number) => {
    const ds = dateStr(d)
    if (selected.includes(ds)) {
      onChange(selected.filter((x) => x !== ds))
    } else {
      onChange([...selected, ds].sort())
    }
  }

  const prevMonth = () => {
    if (month === 0) { setYear((y) => y - 1); setMonth(11) }
    else setMonth((m) => m - 1)
  }
  const nextMonth = () => {
    if (month === 11) { setYear((y) => y + 1); setMonth(0) }
    else setMonth((m) => m + 1)
  }

  const DOW = ['日', '月', '火', '水', '木', '金', '土']

  return (
    <div className="holiday-cal">
      <div className="holiday-cal-header">
        <button type="button" onClick={prevMonth}>‹</button>
        <span>{year}年 {month + 1}月</span>
        <button type="button" onClick={nextMonth}>›</button>
      </div>
      <div className="holiday-cal-grid">
        {DOW.map((d, i) => (
          <div key={d} className={`holiday-cal-dow${i === 0 ? ' sun' : i === 6 ? ' sat' : ''}`}>{d}</div>
        ))}
        {Array.from({ length: firstDow }, (_, i) => (
          <div key={`e${i}`} className="holiday-cal-day empty" />
        ))}
        {Array.from({ length: daysInMonth }, (_, i) => {
          const d = i + 1
          const ds = dateStr(d)
          const dow = (firstDow + i) % 7
          const cls = [
            'holiday-cal-day',
            dow === 0 ? 'sun' : dow === 6 ? 'sat' : '',
            selected.includes(ds) ? 'selected' : '',
            ds === todayStr ? 'today' : '',
          ].filter(Boolean).join(' ')
          return (
            <div key={d} className={cls} onClick={() => toggle(d)}>{d}</div>
          )
        })}
      </div>
      {selected.length > 0 && (
        <div className="holiday-cal-badges">
          {selected.map((h) => (
            <span key={h} className="badge warn" style={{ cursor: 'pointer', fontSize: '12px' }} onClick={() => onChange(selected.filter((x) => x !== h))}>
              {h.replace(/^\d{4}-/, '').replace('-', '/')} ×
            </span>
          ))}
        </div>
      )}
    </div>
  )
}

/** Inline two-month calendar for picking a date range (start–end). */
const DateRangePicker = ({
  startDate, endDate, onStartChange, onEndChange, label,
}: {
  startDate: string; endDate: string;
  onStartChange: (d: string) => void; onEndChange: (d: string) => void;
  label?: string;
}) => {
  const today = new Date()
  const [year, setYear] = useState(() => {
    if (startDate) { const [y] = startDate.split('-'); return Number(y) }
    return today.getFullYear()
  })
  const [month, setMonth] = useState(() => {
    if (startDate) { const parts = startDate.split('-'); return Number(parts[1]) - 1 }
    return today.getMonth()
  })

  const pad2 = (n: number) => String(n).padStart(2, '0')
  const ds = (y: number, m: number, d: number) => `${y}-${pad2(m + 1)}-${pad2(d)}`
  const todayStr = ds(today.getFullYear(), today.getMonth(), today.getDate())

  const handleClick = (dateStr: string) => {
    if (!startDate || (startDate && endDate)) {
      onStartChange(dateStr)
      onEndChange('')
    } else {
      if (dateStr < startDate) {
        onEndChange(startDate)
        onStartChange(dateStr)
      } else if (dateStr === startDate) {
        onStartChange('')
      } else {
        onEndChange(dateStr)
      }
    }
  }

  const prevMonth = () => {
    if (month === 0) { setYear((y) => y - 1); setMonth(11) } else setMonth((m) => m - 1)
  }
  const nextMonth = () => {
    if (month === 11) { setYear((y) => y + 1); setMonth(0) } else setMonth((m) => m + 1)
  }

  const DOW = ['日', '月', '火', '水', '木', '金', '土']
  const m2 = month === 11 ? 0 : month + 1
  const y2 = month === 11 ? year + 1 : year

  const renderMonth = (yr: number, mo: number) => {
    const dim = new Date(yr, mo + 1, 0).getDate()
    const fdow = new Date(yr, mo, 1).getDay()
    return (
      <div className="range-cal-month">
        <div className="range-cal-month-label">{yr}年 {mo + 1}月</div>
        <div className="range-cal-grid">
          {DOW.map((d, i) => (
            <div key={d} className={`range-cal-dow${i === 0 ? ' sun' : i === 6 ? ' sat' : ''}`}>{d}</div>
          ))}
          {Array.from({ length: fdow }, (_, i) => (
            <div key={`e${i}`} className="range-cal-day empty" />
          ))}
          {Array.from({ length: dim }, (_, i) => {
            const d = i + 1
            const dateStr = ds(yr, mo, d)
            const dow = (fdow + i) % 7
            const inRange = startDate && endDate && dateStr >= startDate && dateStr <= endDate
            const cls = [
              'range-cal-day',
              dow === 0 ? 'sun' : dow === 6 ? 'sat' : '',
              inRange ? 'in-range' : '',
              dateStr === startDate ? 'range-start' : '',
              dateStr === endDate ? 'range-end' : '',
              dateStr === todayStr ? 'today' : '',
            ].filter(Boolean).join(' ')
            return (
              <div key={d} className={cls} onClick={() => handleClick(dateStr)}>{d}</div>
            )
          })}
        </div>
      </div>
    )
  }

  return (
    <div className="range-cal">
      {label && <div className="range-cal-label">{label}</div>}
      <div className="range-cal-nav">
        <button type="button" onClick={prevMonth}>‹</button>
        <span style={{ flex: 1 }} />
        <button type="button" onClick={nextMonth}>›</button>
      </div>
      <div className="range-cal-months">
        {renderMonth(year, month)}
        {renderMonth(y2, m2)}
      </div>
      <div className="range-cal-footer">
        {startDate ? (
          <span className="badge ok" style={{ cursor: 'pointer', fontSize: '12px' }} onClick={() => { onStartChange(''); onEndChange('') }}>
            開始: {startDate} ×
          </span>
        ) : (
          <span className="muted" style={{ fontSize: '12px' }}>クリックで開始日を選択</span>
        )}
        {endDate ? (
          <span className="badge ok" style={{ cursor: 'pointer', fontSize: '12px' }} onClick={() => onEndChange('')}>
            終了: {endDate} ×
          </span>
        ) : startDate ? (
          <span className="muted" style={{ fontSize: '12px' }}>クリックで終了日を選択</span>
        ) : null}
      </div>
    </div>
  )
}

const HomePage = () => {
  const { classroomId = '' } = useParams()
  const navigate = useNavigate()
  const location = useLocation()
  const [unlocked, setUnlocked] = useState(false)
  const [showNewSessionForm, setShowNewSessionForm] = useState(false)
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
  const [constraintPersonAType, setConstraintPersonAType] = useState<PairConstraintPersonType>('teacher')
  const [constraintPersonAId, setConstraintPersonAId] = useState('')
  const [constraintPersonBType, setConstraintPersonBType] = useState<PairConstraintPersonType>('student')
  const [constraintPersonBId, setConstraintPersonBId] = useState('')
  const [constraintType, setConstraintType] = useState<ConstraintType>('incompatible')
  const [regularTeacherId, setRegularTeacherId] = useState('')
  const [regularStudent1Id, setRegularStudent1Id] = useState('')
  const [regularStudent2Id, setRegularStudent2Id] = useState('')
  const [regularSubject, setRegularSubject] = useState('')
  const [regularStudentSubjects, setRegularStudentSubjects] = useState<Record<string, string>>({})
  const [regularDayOfWeek, setRegularDayOfWeek] = useState('')
  const [regularSlotNumber, setRegularSlotNumber] = useState('')
  // Group lesson form state
  const [groupTeacherId, setGroupTeacherId] = useState('')
  const [groupSubject, setGroupSubject] = useState('')
  const [groupDayOfWeek, setGroupDayOfWeek] = useState('')
  const [, setGroupSlotNumber] = useState('')  // kept setter for reset calls
  const [groupStudentIds, setGroupStudentIds] = useState<string[]>([])
  const [editingGroupLessonId, setEditingGroupLessonId] = useState<string | null>(null)
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

  // Editing state for regular lessons
  const [editingRegularLessonId, setEditingRegularLessonId] = useState<string | null>(null)

  // Editing state for pair constraints
  const [editingConstraintId, setEditingConstraintId] = useState<string | null>(null)

  useEffect(() => {
    if (import.meta.env.DEV) {
      navigate(`/c/${classroomId}/boot`, { replace: true })
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
    const unsub1 = watchSessionsList(classroomId, (items) => setSessions(items))
    const unsub2 = watchMasterData(classroomId, (md) => {
      if (md) {
        setMasterData(md)
      } else {
        const empty = emptyMasterData()
        saveMasterData(classroomId, empty).catch(() => {})
        setMasterData(empty)
      }
    })
    return () => { unsub1(); unsub2() }
  }, [unlocked, classroomId])

  // --- Save and close: create backup, then navigate to classroom select ---
  const changeLogRef = useRef<Set<string>>(new Set())
  const handleSaveAndClose = async () => {
    if (!classroomId) return
    if (changeLogRef.current.size > 0) {
      try {
        await createBackup(classroomId, 'auto', [...changeLogRef.current])
        await cleanupOldBackups(classroomId, 30)
      } catch (e) {
        console.warn('[SaveAndClose] Backup failed:', e)
      }
    }
    navigate('/')
  }

  // --- Master data helpers ---
  const updateMaster = async (updater: (current: MasterData) => MasterData): Promise<void> => {
    if (!masterData) return
    const next = updater(masterData)
    setMasterData(next)
    await saveMasterData(classroomId, next)
  }

  const addManager = async (): Promise<void> => {
    if (!managerName.trim()) return
    const manager: Manager = { id: createId(), name: managerName.trim(), email: managerEmail.trim() }
    await updateMaster((c) => ({ ...c, managers: [...(c.managers ?? []), manager] }))
    changeLogRef.current.add('マネージャー追加')
    setManagerName(''); setManagerEmail('')
  }

  const addTeacher = async (): Promise<void> => {
    if (!teacherName.trim()) return
    const teacher: Teacher = { id: createId(), name: teacherName.trim(), email: teacherEmail.trim(), subjects: teacherSubjects, memo: teacherMemo.trim() }
    await updateMaster((c) => ({ ...c, teachers: [...c.teachers, teacher] }))
    changeLogRef.current.add('講師追加')
    setTeacherName(''); setTeacherEmail(''); setTeacherSubjects([]); setTeacherMemo('')
  }

  const addStudent = async (): Promise<void> => {
    if (!studentName.trim()) return
    const student: Student = {
      id: createId(), name: studentName.trim(), email: studentEmail.trim(), grade: studentGrade.trim(),
      subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo: '', submittedAt: 0,
    }
    await updateMaster((c) => ({ ...c, students: [...c.students, student] }))
    changeLogRef.current.add('生徒追加')
    setStudentName(''); setStudentEmail(''); setStudentGrade('')
  }

  const upsertConstraint = async (): Promise<void> => {
    if (!constraintPersonAId || !constraintPersonBId || !masterData) return
    if (constraintPersonAId === constraintPersonBId) { alert('同じ人物は選択できません。'); return }
    const nc: PairConstraint = {
      id: createId(),
      personAId: constraintPersonAId,
      personBId: constraintPersonBId,
      personAType: constraintPersonAType,
      personBType: constraintPersonBType,
      type: constraintType,
    }
    await updateMaster((c) => {
      const filtered = c.constraints.filter((i) =>
        !((i.personAId === constraintPersonAId && i.personBId === constraintPersonBId) ||
          (i.personAId === constraintPersonBId && i.personBId === constraintPersonAId)),
      )
      return { ...c, constraints: [...filtered, nc] }
    })
    changeLogRef.current.add('制約変更')
  }

  const addRegularLesson = async (): Promise<void> => {
    const studentIds = [regularStudent1Id, regularStudent2Id].filter(Boolean)
    if (!regularTeacherId || studentIds.length === 0 || !regularDayOfWeek || !regularSlotNumber) return
    // Require all students to have a subject set
    const allHaveSubject = studentIds.every((sid) => regularStudentSubjects[sid])
    if (!allHaveSubject) return
    const defaultSubject = regularStudentSubjects[studentIds[0]] ?? regularSubject
    const nl: RegularLesson = {
      id: createId(), teacherId: regularTeacherId, studentIds, subject: defaultSubject,
      studentSubjects: { ...regularStudentSubjects },
      dayOfWeek: Number.parseInt(regularDayOfWeek, 10), slotNumber: Number.parseInt(regularSlotNumber, 10),
    }
    await updateMaster((c) => ({ ...c, regularLessons: [...c.regularLessons, nl] }))
    changeLogRef.current.add('通常授業追加')
    setRegularTeacherId(''); setRegularStudent1Id(''); setRegularStudent2Id('')
    setRegularSubject(''); setRegularStudentSubjects({}); setRegularDayOfWeek(''); setRegularSlotNumber('')
  }

  const startEditManager = (m: Manager): void => {
    setEditingManagerId(m.id); setEditManagerName(m.name); setEditManagerEmail(m.email || '')
  }
  const saveEditManager = async (): Promise<void> => {
    if (!editingManagerId || !editManagerName.trim()) return
    await updateMaster((c) => ({ ...c, managers: (c.managers ?? []).map((m) => m.id === editingManagerId ? { ...m, name: editManagerName.trim(), email: editManagerEmail.trim() } : m) }))
    changeLogRef.current.add('マネージャー編集')
    setEditingManagerId(null)
  }

  const startEditTeacher = (t: Teacher): void => {
    setEditingTeacherId(t.id); setEditTeacherName(t.name); setEditTeacherEmail(t.email || '')
    setEditTeacherSubjects([...t.subjects]); setEditTeacherMemo(t.memo || '')
  }
  const saveEditTeacher = async (): Promise<void> => {
    if (!editingTeacherId || !editTeacherName.trim()) return
    await updateMaster((c) => ({ ...c, teachers: c.teachers.map((t) => t.id === editingTeacherId ? { ...t, name: editTeacherName.trim(), email: editTeacherEmail.trim(), subjects: editTeacherSubjects, memo: editTeacherMemo.trim() } : t) }))
    changeLogRef.current.add('講師編集')
    setEditingTeacherId(null)
  }

  const startEditStudent = (s: Student): void => {
    setEditingStudentId(s.id); setEditStudentName(s.name); setEditStudentEmail(s.email || '')
    setEditStudentGrade(s.grade || '')
  }
  const saveEditStudent = async (): Promise<void> => {
    if (!editingStudentId || !editStudentName.trim()) return
    await updateMaster((c) => ({ ...c, students: c.students.map((s) => s.id === editingStudentId ? { ...s, name: editStudentName.trim(), email: editStudentEmail.trim(), grade: editStudentGrade.trim() } : s) }))
    changeLogRef.current.add('生徒編集')
    setEditingStudentId(null)
  }

  const removeManager = async (managerId: string): Promise<void> => {
    if (!window.confirm('このマネージャーを削除しますか？')) return
    await updateMaster((c) => ({ ...c, managers: (c.managers ?? []).filter((m) => m.id !== managerId) }))
    changeLogRef.current.add('マネージャー削除')
  }

  const removeTeacher = async (teacherId: string): Promise<void> => {
    if (!window.confirm('この講師を削除しますか？')) return
    await updateMaster((c) => ({ ...c, teachers: c.teachers.filter((t) => t.id !== teacherId) }))
    changeLogRef.current.add('講師削除')
  }

  const removeStudent = async (studentId: string): Promise<void> => {
    if (!window.confirm('この生徒を削除しますか？')) return
    await updateMaster((c) => ({ ...c, students: c.students.filter((s) => s.id !== studentId) }))
    changeLogRef.current.add('生徒削除')
  }

  const removeConstraint = async (constraintId: string): Promise<void> => {
    await updateMaster((c) => ({ ...c, constraints: c.constraints.filter((x) => x.id !== constraintId) }))
    changeLogRef.current.add('制約削除')
  }

  const removeRegularLesson = async (lessonId: string): Promise<void> => {
    await updateMaster((c) => ({ ...c, regularLessons: c.regularLessons.filter((l) => l.id !== lessonId) }))
    changeLogRef.current.add('通常授業削除')
  }

  // Edit regular lesson: populate form fields with existing data
  const startEditRegularLesson = (l: RegularLesson): void => {
    setEditingRegularLessonId(l.id)
    setRegularTeacherId(l.teacherId)
    setRegularStudent1Id(l.studentIds[0] ?? '')
    setRegularStudent2Id(l.studentIds[1] ?? '')
    setRegularSubject(l.subject)
    setRegularStudentSubjects(l.studentSubjects ?? {})
    setRegularDayOfWeek(String(l.dayOfWeek))
    setRegularSlotNumber(String(l.slotNumber))
  }

  const saveEditRegularLesson = async (): Promise<void> => {
    if (!editingRegularLessonId) return
    const studentIds = [regularStudent1Id, regularStudent2Id].filter(Boolean)
    if (!regularTeacherId || studentIds.length === 0 || !regularDayOfWeek || !regularSlotNumber) return
    const allHaveSubject = studentIds.every((sid) => regularStudentSubjects[sid])
    if (!allHaveSubject) return
    const defaultSubject = regularStudentSubjects[studentIds[0]] ?? regularSubject
    await updateMaster((c) => ({
      ...c,
      regularLessons: c.regularLessons.map((l) =>
        l.id === editingRegularLessonId
          ? { ...l, teacherId: regularTeacherId, studentIds, subject: defaultSubject, studentSubjects: { ...regularStudentSubjects }, dayOfWeek: Number.parseInt(regularDayOfWeek, 10), slotNumber: Number.parseInt(regularSlotNumber, 10) }
          : l,
      ),
    }))
    changeLogRef.current.add('通常授業編集')
    setEditingRegularLessonId(null)
    setRegularTeacherId(''); setRegularStudent1Id(''); setRegularStudent2Id('')
    setRegularSubject(''); setRegularStudentSubjects({}); setRegularDayOfWeek(''); setRegularSlotNumber('')
  }

  const cancelEditRegularLesson = (): void => {
    setEditingRegularLessonId(null)
    setRegularTeacherId(''); setRegularStudent1Id(''); setRegularStudent2Id('')
    setRegularSubject(''); setRegularStudentSubjects({}); setRegularDayOfWeek(''); setRegularSlotNumber('')
  }

  // Group lesson CRUD
  const addGroupLesson = async (): Promise<void> => {
    if (!groupTeacherId || !groupSubject || !groupDayOfWeek || groupStudentIds.length === 0) return
    const nl: GroupLesson = {
      id: createId(), teacherId: groupTeacherId, studentIds: [...groupStudentIds], subject: groupSubject,
      dayOfWeek: Number.parseInt(groupDayOfWeek, 10), slotNumber: 0,
    }
    await updateMaster((c) => ({ ...c, groupLessons: [...(c.groupLessons ?? []), nl] }))
    changeLogRef.current.add('集団授業追加')
    setGroupTeacherId(''); setGroupSubject(''); setGroupDayOfWeek(''); setGroupSlotNumber(''); setGroupStudentIds([])
  }

  const removeGroupLesson = async (lessonId: string): Promise<void> => {
    await updateMaster((c) => ({ ...c, groupLessons: (c.groupLessons ?? []).filter((l) => l.id !== lessonId) }))
    changeLogRef.current.add('集団授業削除')
  }

  const startEditGroupLesson = (l: GroupLesson): void => {
    setEditingGroupLessonId(l.id)
    setGroupTeacherId(l.teacherId)
    setGroupSubject(l.subject)
    setGroupDayOfWeek(String(l.dayOfWeek))
    setGroupStudentIds([...l.studentIds])
  }

  const saveEditGroupLesson = async (): Promise<void> => {
    if (!editingGroupLessonId || !groupTeacherId || !groupSubject || !groupDayOfWeek || groupStudentIds.length === 0) return
    await updateMaster((c) => ({
      ...c,
      groupLessons: (c.groupLessons ?? []).map((l) =>
        l.id === editingGroupLessonId
          ? { ...l, teacherId: groupTeacherId, studentIds: [...groupStudentIds], subject: groupSubject, dayOfWeek: Number.parseInt(groupDayOfWeek, 10), slotNumber: 0 }
          : l,
      ),
    }))
    changeLogRef.current.add('集団授業編集')
    setEditingGroupLessonId(null)
    setGroupTeacherId(''); setGroupSubject(''); setGroupDayOfWeek(''); setGroupSlotNumber(''); setGroupStudentIds([])
  }

  const cancelEditGroupLesson = (): void => {
    setEditingGroupLessonId(null)
    setGroupTeacherId(''); setGroupSubject(''); setGroupDayOfWeek(''); setGroupSlotNumber(''); setGroupStudentIds([])
  }

  // Edit pair constraint: populate form
  const startEditConstraint = (c: PairConstraint): void => {
    setEditingConstraintId(c.id)
    setConstraintPersonAType(c.personAType)
    setConstraintPersonAId(c.personAId)
    setConstraintPersonBType(c.personBType)
    setConstraintPersonBId(c.personBId)
    setConstraintType(c.type)
  }

  const saveEditConstraint = async (): Promise<void> => {
    if (!editingConstraintId || !constraintPersonAId || !constraintPersonBId) return
    await updateMaster((c) => ({
      ...c,
      constraints: c.constraints.map((item) =>
        item.id === editingConstraintId
          ? { ...item, personAId: constraintPersonAId, personBId: constraintPersonBId, personAType: constraintPersonAType, personBType: constraintPersonBType, type: constraintType }
          : item,
      ),
    }))
    changeLogRef.current.add('制約編集')
    setEditingConstraintId(null)
    setConstraintPersonAId(''); setConstraintPersonBId('')
  }

  const cancelEditConstraint = (): void => {
    setEditingConstraintId(null)
    setConstraintPersonAId(''); setConstraintPersonBId('')
  }

  // --- Bulk delete ---
  const handleBulkDelete = async (): Promise<void> => {
    if (!masterData) return
    const pw = window.prompt('管理データを一括削除します。\nパスワードを入力してください:', 'admin1234')
    if (pw === null) return
    if (pw !== adminPassword) {
      alert('パスワードが違います。')
      return
    }
    if (!window.confirm('本当にすべての管理データ（マネージャー・講師・生徒・制約・通常授業）を削除しますか？\nこの操作は取り消せません。')) return
    await updateMaster(() => emptyMasterData())
    changeLogRef.current.add('管理データ一括削除')
    alert('管理データを一括削除しました。')
  }

  // --- Excel (operates on master data) ---
  const downloadTemplate = (): void => {
    // Realistic sample data: 1 manager, 10 teachers, 30 students with regular & group lessons
    const sampleManagers = [
      ['山田 太郎', 'yamada@example.com'],
    ]
    const sampleTeachers = [
      ['田中 一郎', '高数,高英', '', ''],
      ['佐藤 花子', '高英,中国', '', ''],
      ['鈴木 健太', '高数,高理', '', ''],
      ['高橋 美咲', '中英,中数', '', ''],
      ['伊藤 大輔', '高英,高社', '', ''],
      ['渡辺 さくら', '高国,中英', '', ''],
      ['山本 翔太', '高数,高理,高IT', '', ''],
      ['中村 あおい', '中数,中英,中国', '', ''],
      ['小林 拓海', '高英,高数', '', ''],
      ['加藤 凛', '高国,高社,中理', '', ''],
    ]
    const sampleStudents = [
      ['青木 太郎', '中3', '', ''], ['石井 花', '中2', '', ''], ['上田 陽介', '高1', '', ''],
      ['遠藤 美月', '高2', '', ''], ['大野 蓮', '中3', '', ''], ['岡田 結衣', '中1', '', ''],
      ['川口 翼', '高3', '', ''], ['北村 陽菜', '高1', '', ''], ['木村 悠真', '中2', '', ''],
      ['小松 あかり', '高2', '', ''], ['坂本 海斗', '中3', '', ''], ['佐々木 莉子', '中1', '', ''],
      ['島田 颯太', '高1', '', ''], ['杉山 美優', '高3', '', ''], ['関口 大和', '中2', '', ''],
      ['高木 七海', '高2', '', ''], ['竹内 蒼空', '中3', '', ''], ['田村 琴音', '中1', '', ''],
      ['中島 陸', '高1', '', ''], ['西田 ひかり', '高3', '', ''], ['野口 春樹', '中2', '', ''],
      ['橋本 凪', '高2', '', ''], ['林 瑠奈', '中3', '', ''], ['平野 壮太', '中1', '', ''],
      ['藤田 真白', '高1', '', ''], ['古川 湊', '高3', '', ''], ['松田 彩花', '中2', '', ''],
      ['三浦 律', '高2', '', ''], ['宮崎 詩', '中3', '', ''], ['森 大智', '中1', '', ''],
    ]
    const sampleConstraints = [
      ['講師', '田中 一郎', '生徒', '石井 花', '不可'],
      ['生徒', '青木 太郎', '生徒', '上田 陽介', '不可'],
    ]
    const sampleRegularLessons = [
      ['田中 一郎', '青木 太郎', '数', '大野 蓮', '数', '月', '1'],
      ['佐藤 花子', '石井 花', '英', '', '', '月', '2'],
      ['鈴木 健太', '上田 陽介', '数', '北村 陽菜', '数', '火', '1'],
      ['高橋 美咲', '岡田 結衣', '英', '木村 悠真', '数', '火', '2'],
      ['伊藤 大輔', '遠藤 美月', '英', '', '', '水', '1'],
      ['渡辺 さくら', '川口 翼', '国', '杉山 美優', '国', '水', '2'],
      ['山本 翔太', '坂本 海斗', '数', '', '', '木', '1'],
      ['中村 あおい', '佐々木 莉子', '数', '田村 琴音', '英', '木', '2'],
      ['小林 拓海', '島田 颯太', '英', '中島 陸', '数', '金', '1'],
      ['加藤 凛', '竹内 蒼空', '国', '', '', '金', '2'],
    ]
    const sampleGroupLessons = [
      ['田中 一郎', '数', '青木 太郎, 大野 蓮, 坂本 海斗, 竹内 蒼空, 林 瑠奈, 宮崎 詩', '土'],
      ['佐藤 花子', '英', '石井 花, 木村 悠真, 関口 大都, 野口 春樹, 松田 彩花', '土'],
    ]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', 'メール'], ...sampleManagers]), 'マネージャー')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '担当科目(カンマ区切り: ' + ALL_TEACHER_SUBJECTS.join(',') + ')', 'メモ', 'メール'], ...sampleTeachers]), '講師')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '学年', 'メモ', 'メール'], ...sampleStudents]), '生徒')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['人物A種別', '人物A名', '人物B種別', '人物B名', '種別'], ...sampleConstraints]), '制約')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒1名', '生徒1科目', '生徒2名', '生徒2科目', '曜日', '時限'], ...sampleRegularLessons]), '通常授業')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '科目', '生徒名(カンマ区切り)', '曜日'], ...sampleGroupLessons]), '集団授業')
    XLSX.writeFile(wb, 'テンプレート.xlsx')
  }

  const exportData = (): void => {
    if (!masterData) return
    const md = masterData
    const teacherRows = md.teachers.map((t) => [t.name, t.subjects.join(', '), t.memo, t.email ?? ''])
    const studentRows = md.students.map((s) => [s.name, s.grade, s.memo, s.email ?? ''])
    const findPersonName = (id: string, pType: string): string => {
      if (pType === 'teacher') return md.teachers.find((t) => t.id === id)?.name ?? id
      return md.students.find((s) => s.id === id)?.name ?? id
    }
    const constraintRows = md.constraints.map((c) => [
      c.personAType === 'teacher' ? '講師' : '生徒',
      findPersonName(c.personAId, c.personAType),
      c.personBType === 'teacher' ? '講師' : '生徒',
      findPersonName(c.personBId, c.personBType),
      c.type === 'incompatible' ? '不可' : '推奨',
    ])
    const dayNames = ['日', '月', '火', '水', '木', '金', '土']
    const regularLessonRows = md.regularLessons.map((l) => {
      const s1Name = l.studentIds[0] ? (md.students.find((s) => s.id === l.studentIds[0])?.name ?? l.studentIds[0]) : ''
      const s1Subj = l.studentIds[0] ? (l.studentSubjects?.[l.studentIds[0]] ?? l.subject) : ''
      const s2Name = l.studentIds[1] ? (md.students.find((s) => s.id === l.studentIds[1])?.name ?? l.studentIds[1]) : ''
      const s2Subj = l.studentIds[1] ? (l.studentSubjects?.[l.studentIds[1]] ?? l.subject) : ''
      return [
        md.teachers.find((t) => t.id === l.teacherId)?.name ?? l.teacherId,
        s1Name, s1Subj, s2Name, s2Subj,
        dayNames[l.dayOfWeek] ?? '', l.slotNumber,
      ]
    })
    const managerRows = (md.managers ?? []).map((m) => [m.name, m.email ?? ''])
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', 'メール'], ...managerRows]), 'マネージャー')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '担当科目', 'メモ', 'メール'], ...teacherRows]), '講師')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['名前', '学年', 'メモ', 'メール'], ...studentRows]), '生徒')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['人物A種別', '人物A名', '人物B種別', '人物B名', '種別'], ...constraintRows]), '制約')
    const groupLessonRows = (md.groupLessons ?? []).map((gl) => [
      md.teachers.find((t) => t.id === gl.teacherId)?.name ?? gl.teacherId,
      gl.subject,
      gl.studentIds.map((sid) => md.students.find((s) => s.id === sid)?.name ?? sid).join(', '),
      dayNames[gl.dayOfWeek] ?? '',
    ])
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '生徒1名', '生徒1科目', '生徒2名', '生徒2科目', '曜日', '時限'], ...regularLessonRows]), '通常授業')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['講師名', '科目', '生徒名(カンマ区切り)', '曜日'], ...groupLessonRows]), '集団授業')
    XLSX.writeFile(wb, '管理データ.xlsx')
  }

  const handleFileImport = async (file: File): Promise<void> => {
    if (!masterData) return
    const md = masterData
    const buf = await file.arrayBuffer()
    const wb = XLSX.read(buf, { type: 'array' })

    const importedManagers: Manager[] = []
    const importedTeachers: Teacher[] = []
    const importedStudents: Student[] = []
    const importedConstraints: PairConstraint[] = []
    const importedRegularLessons: RegularLesson[] = []
    const importedGroupLessons: GroupLesson[] = []

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

    const managerWs = wb.Sheets['マネージャー']
    if (managerWs) {
      const rows = XLSX.utils.sheet_to_json(managerWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const name = String(row?.[0] ?? '').trim()
        if (!name) continue
        const email = String(row?.[1] ?? '').trim()
        if ((md.managers ?? []).some((m) => m.name === name)) continue
        importedManagers.push({ id: createId(), name, email })
      }
    }

    const teacherWs = wb.Sheets['講師']
    if (teacherWs) {
      const rows = XLSX.utils.sheet_to_json(teacherWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const name = String(row?.[0] ?? '').trim()
        if (!name) continue
        const subjects = String(row?.[1] ?? '').split(/[、,]/).map((s) => s.trim()).filter((s) => ALL_TEACHER_SUBJECTS.includes(s) || FIXED_SUBJECTS.includes(s))
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
        const memo = String(row?.[2] ?? '').trim()
        const email = String(row?.[3] ?? '').trim()
        if (md.students.some((s) => s.name === name)) continue
        importedStudents.push({ id: createId(), name, email, grade, subjects: [], subjectSlots: {}, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo, submittedAt: 0 })
      }
    }

    const constraintWs = wb.Sheets['制約']
    if (constraintWs) {
      const rows = XLSX.utils.sheet_to_json(constraintWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const aTypeStr = String(row?.[0] ?? '').trim()
        const aName = String(row?.[1] ?? '').trim()
        const bTypeStr = String(row?.[2] ?? '').trim()
        const bName = String(row?.[3] ?? '').trim()
        const ts = String(row?.[4] ?? '').trim()
        const personAType: PairConstraintPersonType = aTypeStr === '生徒' ? 'student' : 'teacher'
        const personBType: PairConstraintPersonType = bTypeStr === '生徒' ? 'student' : 'teacher'
        const aid = personAType === 'teacher' ? findTeacherId(aName) : findStudentId(aName)
        const bid = personBType === 'teacher' ? findTeacherId(bName) : findStudentId(bName)
        if (!aid || !bid) continue
        const type: ConstraintType = ts === '推奨' ? 'recommended' : 'incompatible'
        if (md.constraints.some((c) =>
          (c.personAId === aid && c.personBId === bid) || (c.personAId === bid && c.personBId === aid),
        )) continue
        importedConstraints.push({ id: createId(), personAId: aid, personBId: bid, personAType, personBType, type })
      }
    }

    const dayNameMap: Record<string, number> = { '日': 0, '月': 1, '火': 2, '水': 3, '木': 4, '金': 5, '土': 6 }
    const regularWs = wb.Sheets['通常授業']
    if (regularWs) {
      const rows = XLSX.utils.sheet_to_json(regularWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const tName = String(row?.[0] ?? '').trim(); const s1 = String(row?.[1] ?? '').trim(); const subj1 = String(row?.[2] ?? '').trim()
        const s2 = String(row?.[3] ?? '').trim(); const subj2 = String(row?.[4] ?? '').trim()
        const dayStr = String(row?.[5] ?? '').trim(); const slotNum = Number(row?.[6])
        const tid = findTeacherId(tName)
        if (!tid || !subj1) continue
        const sid1 = findStudentId(s1); const sid2 = findStudentId(s2)
        const sids = [sid1, sid2].filter(Boolean) as string[]
        if (sids.length === 0) continue
        const dow = dayNameMap[dayStr]
        if (dow === undefined || Number.isNaN(slotNum) || slotNum < 1) continue
        const studentSubjects: Record<string, string> = {}
        if (sid1) studentSubjects[sid1] = subj1
        if (sid2 && subj2) studentSubjects[sid2] = subj2
        else if (sid2) studentSubjects[sid2] = subj1
        const subject = subj1
        // Dedup: skip if same regular lesson already exists in master data
        const sortedSids = [...sids].sort()
        const isDup = md.regularLessons.some((existing) =>
          existing.teacherId === tid &&
          existing.dayOfWeek === dow &&
          existing.slotNumber === slotNum &&
          existing.studentIds.length === sortedSids.length &&
          [...existing.studentIds].sort().every((id, j) => id === sortedSids[j]),
        )
        if (isDup) continue
        importedRegularLessons.push({ id: createId(), teacherId: tid, studentIds: sids, subject, studentSubjects, dayOfWeek: dow, slotNumber: slotNum })
      }
    }

    const groupWs = wb.Sheets['集団授業']
    if (groupWs) {
      const rows = XLSX.utils.sheet_to_json(groupWs, { header: 1 }) as unknown as unknown[][]
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i]
        const tName = String(row?.[0] ?? '').trim()
        const subject = String(row?.[1] ?? '').trim()
        const studentsStr = String(row?.[2] ?? '').trim()
        const dayStr = String(row?.[3] ?? '').trim()
        const tid = findTeacherId(tName)
        if (!tid || !subject) continue
        const dow = dayNameMap[dayStr]
        if (dow === undefined) continue
        const studentNames = studentsStr.split(/[、,]/).map((s) => s.trim()).filter(Boolean)
        const sids = studentNames.map((n) => findStudentId(n)).filter(Boolean) as string[]
        if (sids.length === 0) continue
        // Dedup: skip if same group lesson already exists
        const sortedSids = [...sids].sort()
        const isDup = (md.groupLessons ?? []).some((existing) =>
          existing.teacherId === tid &&
          existing.dayOfWeek === dow &&
          existing.studentIds.length === sortedSids.length &&
          [...existing.studentIds].sort().every((id, j) => id === sortedSids[j]),
        )
        if (isDup) continue
        importedGroupLessons.push({ id: createId(), teacherId: tid, studentIds: sids, subject, dayOfWeek: dow, slotNumber: 0 })
      }
    }

    // --- Validation: check for inconsistencies ---
    const validationWarnings: string[] = []
    // All teachers (existing + imported)
    const allTeachers = [...md.teachers, ...importedTeachers]
    const allStudents = [...md.students, ...importedStudents]
    // Check regular lessons — remove invalid ones but keep valid ones
    const validRegularLessons = importedRegularLessons.filter((rl) => {
      const teacher = allTeachers.find((t) => t.id === rl.teacherId)
      const dayNames2 = ['日', '月', '火', '水', '木', '金', '土']
      let valid = true
      // Check all per-student subjects (and fallback subject)
      const subjectsToCheck = rl.studentSubjects
        ? Object.values(rl.studentSubjects)
        : [rl.subject]
      for (const subj of subjectsToCheck) {
        if (teacher && !teacherHasSubject(teacher.subjects, subj)) {
          validationWarnings.push(`通常授業スキップ: ${teacher.name} の担当科目に「${subj}」がありません（${dayNames2[rl.dayOfWeek]}曜${rl.slotNumber}限）`)
          valid = false
        }
      }
      for (const sid of rl.studentIds) {
        if (!allStudents.find((s) => s.id === sid)) {
          validationWarnings.push(`通常授業スキップ: 生徒ID「${sid}」が見つかりません（${dayNames2[rl.dayOfWeek]}曜${rl.slotNumber}限）`)
          valid = false
        }
      }
      return valid
    })
    // Check group lessons
    const validGroupLessons = importedGroupLessons.filter((gl) => {
      const teacher = allTeachers.find((t) => t.id === gl.teacherId)
      const dayNames2 = ['日', '月', '火', '水', '木', '金', '土']
      let valid = true
      if (teacher && !teacherHasSubject(teacher.subjects, gl.subject)) {
        validationWarnings.push(`集団授業スキップ: ${teacher.name} の担当科目に「${gl.subject}」がありません（${dayNames2[gl.dayOfWeek]}曜${gl.slotNumber}限）`)
        valid = false
      }
      for (const sid of gl.studentIds) {
        if (!allStudents.find((s) => s.id === sid)) {
          validationWarnings.push(`集団授業スキップ: 生徒ID「${sid}」が見つかりません（${dayNames2[gl.dayOfWeek]}曜${gl.slotNumber}限）`)
          valid = false
        }
      }
      return valid
    })
    // Check constraints reference valid people — remove invalid ones
    const validConstraints = importedConstraints.filter((c) => {
      const aList = c.personAType === 'teacher' ? allTeachers : allStudents
      const bList = c.personBType === 'teacher' ? allTeachers : allStudents
      let valid = true
      if (!aList.some((p) => p.id === c.personAId)) {
        validationWarnings.push(`制約スキップ: 人物A「${c.personAId}」が見つかりません`)
        valid = false
      }
      if (!bList.some((p) => p.id === c.personBId)) {
        validationWarnings.push(`制約スキップ: 人物B「${c.personBId}」が見つかりません`)
        valid = false
      }
      return valid
    })

    const added: string[] = []
    if (importedManagers.length) added.push(`マネージャー${importedManagers.length}名`)
    if (importedTeachers.length) added.push(`講師${importedTeachers.length}名`)
    if (importedStudents.length) added.push(`生徒${importedStudents.length}名`)
    if (validConstraints.length) added.push(`制約${validConstraints.length}件`)
    if (validRegularLessons.length) added.push(`通常授業${validRegularLessons.length}件`)
    if (validGroupLessons.length) added.push(`集団授業${validGroupLessons.length}件`)
    if (added.length === 0 && validationWarnings.length === 0) { alert('新規データがありませんでした（同名は重複スキップ）。'); return }
    if (added.length === 0 && validationWarnings.length > 0) { alert(`⚠️ 以下のデータにエラーがあり、取り込めるデータがありませんでした:\n\n${validationWarnings.join('\n')}`); return }
    const confirmMsg = validationWarnings.length > 0
      ? `以下を取り込みます:\n${added.join(', ')}\n\n⚠️ スキップされたデータ:\n${validationWarnings.join('\n')}\n\nよろしいですか？`
      : `以下を取り込みます:\n${added.join(', ')}\n\nよろしいですか？`
    if (!window.confirm(confirmMsg)) return

    await updateMaster((c) => ({
      ...c,
      managers: [...(c.managers ?? []), ...importedManagers],
      teachers: [...c.teachers, ...importedTeachers],
      students: [...c.students, ...importedStudents],
      constraints: [...c.constraints, ...validConstraints],
      regularLessons: [...c.regularLessons, ...validRegularLessons],
      groupLessons: [...(c.groupLessons ?? []), ...validGroupLessons],
    }))
    changeLogRef.current.add(`ファイル取り込み (${added.join(', ')})`)
    setTimeout(() => alert('取り込み完了！'), 50)
  }

  // --- Session management ---
  const cleanupLegacyDevSession = async (): Promise<void> => {
    const legacyDev = await loadSession(classroomId, 'dev')
    if (!legacyDev) return
    await deleteSession(classroomId, 'dev')
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
    seed.regularLessons = isMendanSession ? [] : masterData.regularLessons
    try {
      const verified = await saveAndVerify(classroomId, id, seed)
      if (!verified) {
        alert('特別講習の作成に失敗しました。Firebaseのセキュリティルールを確認してください。')
      }
    } catch (e) {
      alert(`特別講習の作成に失敗しました:\n${e instanceof Error ? e.message : String(e)}`)
    }
  }

  const openAdmin = (sessionId: string): void => {
    saveAdminPassword(adminPassword)
    navigate(`/c/${classroomId}/admin/${sessionId}`)
  }

  const handleDeleteSession = async (sessionId: string, sessionName: string): Promise<void> => {
    const confirmed = window.confirm(`特別講習「${sessionName || sessionId}」を削除しますか？\nこの操作は元に戻せません。`)
    if (!confirmed) return
    const password = window.prompt('削除パスワードを入力してください:')
    if (password !== adminPassword) {
      alert('パスワードが正しくありません。')
      return
    }
    await deleteSession(classroomId, sessionId)
    changeLogRef.current.add(`セッション削除 (${sessionName || sessionId})`)
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
        <div className="row" style={{ justifyContent: 'space-between', alignItems: 'center' }}>
          <h2 style={{ margin: 0 }}>講習コマ割りアプリ</h2>
          {unlocked && <button className="btn btn-primary" type="button" onClick={() => void handleSaveAndClose()}>保存して閉じる</button>}
        </div>
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
              <div className="row" style={{ justifyContent: 'space-between', alignItems: 'center' }}>
                <h3>特別講習一覧（新しい順）</h3>
                <button className="btn" type="button" onClick={() => setShowNewSessionForm((v) => !v)}>
                  {showNewSessionForm ? '閉じる' : '＋ 追加'}
                </button>
              </div>

              {showNewSessionForm && (
                <div style={{ borderTop: '1px solid #e5e7eb', paddingTop: 12, marginTop: 8, marginBottom: 12 }}>
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
                      <div className="session-create-calendars">
                        <DateRangePicker
                          label="📅 講習期間"
                          startDate={newStartDate} endDate={newEndDate}
                          onStartChange={setNewStartDate} onEndChange={setNewEndDate}
                        />
                        <DateRangePicker
                          label="📝 提出期間 ※この期間のみ希望URLが有効"
                          startDate={newSubmissionStart} endDate={newSubmissionEnd}
                          onStartChange={setNewSubmissionStart} onEndChange={setNewSubmissionEnd}
                        />
                      </div>
                      <div style={{ marginTop: '12px' }}>
                        <span className="muted" style={{ marginBottom: '4px', display: 'block' }}>🚫 休日: (日付をクリックで選択/解除)</span>
                        <HolidayCalendar selected={newHolidays} onChange={setNewHolidays} />
                      </div>
                      <div className="row" style={{ marginTop: '12px' }}>
                        <button className="btn" type="button" onClick={() => void onCreateSession()}>特別講習を作成</button>
                      </div>
                    </>
                  )}
                </div>
              )}

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
                    <button className="btn secondary" type="button" style={{ color: '#dc2626', marginLeft: 'auto' }} onClick={() => void handleBulkDelete()}>管理データ一括削除</button>
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
                      {ALL_TEACHER_SUBJECTS.filter((s) => !teacherSubjects.includes(s)).map((s) => (<option key={s} value={s}>{s}</option>))}
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
                                {ALL_TEACHER_SUBJECTS.filter((s) => !editTeacherSubjects.includes(s)).map((s) => (<option key={s} value={s}>{s}</option>))}
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
                  <div className="row" style={{ flexWrap: 'wrap', gap: 8 }}>
                    <select value={regularTeacherId} onChange={(e) => setRegularTeacherId(e.target.value)}>
                      <option value="">講師を選択</option>
                      {masterData.teachers.map((t) => (<option key={t.id} value={t.id}>{t.name}</option>))}
                    </select>
                    <select value={regularStudent1Id} onChange={(e) => { setRegularStudent1Id(e.target.value); setRegularStudentSubjects((prev) => { const next = { ...prev }; delete next[regularStudent1Id]; return next }) }}>
                      <option value="">生徒1を選択</option>
                      {masterData.students.map((s) => (<option key={s.id} value={s.id} disabled={s.id === regularStudent2Id}>{s.name}</option>))}
                    </select>
                    <select value={regularStudent2Id} onChange={(e) => { setRegularStudent2Id(e.target.value); setRegularStudentSubjects((prev) => { const next = { ...prev }; delete next[regularStudent2Id]; return next }) }}>
                      <option value="">生徒2(任意)</option>
                      {masterData.students.map((s) => (<option key={s.id} value={s.id} disabled={s.id === regularStudent1Id}>{s.name}</option>))}
                    </select>
                    <select value={regularDayOfWeek} onChange={(e) => setRegularDayOfWeek(e.target.value)}>
                      <option value="">曜日を選択</option>
                      <option value="0">日曜</option><option value="1">月曜</option><option value="2">火曜</option>
                      <option value="3">水曜</option><option value="4">木曜</option><option value="5">金曜</option><option value="6">土曜</option>
                    </select>
                    <input type="number" value={regularSlotNumber} onChange={(e) => setRegularSlotNumber(e.target.value)} placeholder="時限番号" min="1" />
                    {editingRegularLessonId ? (
                      <>
                        <button className="btn" type="button" onClick={() => void saveEditRegularLesson()}>更新</button>
                        <button className="btn secondary" type="button" onClick={cancelEditRegularLesson}>キャンセル</button>
                      </>
                    ) : (
                      <button className="btn" type="button" onClick={() => void addRegularLesson()}>追加</button>
                    )}
                  </div>
                  {/* Per-student subject selects */}
                  {(() => {
                    const selectedStudentIds = [regularStudent1Id, regularStudent2Id].filter(Boolean)
                    const teacher = masterData.teachers.find((t) => t.id === regularTeacherId)
                    if (selectedStudentIds.length === 0) return null
                    return (
                      <div style={{ marginTop: 8, display: 'flex', gap: 12, flexWrap: 'wrap' }}>
                        {selectedStudentIds.map((sid, idx) => {
                          const student = masterData.students.find((s) => s.id === sid)
                          if (!student) return null
                          const availableSubjects = teacher
                            ? teachableBaseSubjects(teacher.subjects, student.grade)
                            : [...BASE_SUBJECTS as readonly string[]]
                          return (
                            <label key={sid} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                              <span>{student.name}(生徒{idx + 1}):</span>
                              <select
                                value={regularStudentSubjects[sid] ?? ''}
                                onChange={(e) => setRegularStudentSubjects((prev) => ({ ...prev, [sid]: e.target.value }))}
                              >
                                <option value="">科目を選択</option>
                                {availableSubjects.map((s) => (<option key={s} value={s}>{s}</option>))}
                              </select>
                            </label>
                          )
                        })}
                      </div>
                    )
                  })()}
                  <p className="muted">通常授業は該当する曜日・時限のスロットに最優先で割り当てられます。</p>
                  <table className="table">
                    <thead><tr><th>講師</th><th>生徒1</th><th>科目</th><th>生徒2</th><th>科目</th><th>曜日</th><th>時限</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.regularLessons.map((l) => {
                        const dayNames = ['日', '月', '火', '水', '木', '金', '土']
                        const s1Id = l.studentIds[0]
                        const s2Id = l.studentIds[1]
                        return (
                          <tr key={l.id}>
                            <td>{masterData.teachers.find((t) => t.id === l.teacherId)?.name ?? '-'}</td>
                            <td>{s1Id ? (masterData.students.find((s) => s.id === s1Id)?.name ?? '-') : ''}</td>
                            <td>{s1Id ? (l.studentSubjects?.[s1Id] ?? l.subject) : ''}</td>
                            <td>{s2Id ? (masterData.students.find((s) => s.id === s2Id)?.name ?? '-') : ''}</td>
                            <td>{s2Id ? (l.studentSubjects?.[s2Id] ?? l.subject) : ''}</td>
                            <td>{dayNames[l.dayOfWeek]}曜</td><td>{l.slotNumber}限</td>
                            <td>
                              <button className="btn secondary" type="button" style={{ marginRight: '4px' }} onClick={() => startEditRegularLesson(l)}>編集</button>
                              <button className="btn secondary" type="button" onClick={() => void removeRegularLesson(l.id)}>削除</button>
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>集団授業設定</h3>
                  <div className="row" style={{ flexWrap: 'wrap', gap: 8 }}>
                    <select value={groupTeacherId} onChange={(e) => setGroupTeacherId(e.target.value)}>
                      <option value="">講師を選択</option>
                      {masterData.teachers.map((t) => (<option key={t.id} value={t.id}>{t.name}</option>))}
                    </select>
                    <select value={groupSubject} onChange={(e) => setGroupSubject(e.target.value)}>
                      <option value="">科目を選択</option>
                      {(BASE_SUBJECTS as readonly string[]).map((s) => (<option key={s} value={s}>{s}</option>))}
                    </select>
                    <select value={groupDayOfWeek} onChange={(e) => setGroupDayOfWeek(e.target.value)}>
                      <option value="">曜日を選択</option>
                      <option value="0">日曜</option><option value="1">月曜</option><option value="2">火曜</option>
                      <option value="3">水曜</option><option value="4">木曜</option><option value="5">金曜</option><option value="6">土曜</option>
                    </select>
                    {editingGroupLessonId ? (
                      <>
                        <button className="btn" type="button" onClick={() => void saveEditGroupLesson()}>更新</button>
                        <button className="btn secondary" type="button" onClick={cancelEditGroupLesson}>キャンセル</button>
                      </>
                    ) : (
                      <button className="btn" type="button" onClick={() => void addGroupLesson()}>追加</button>
                    )}
                  </div>
                  <div style={{ marginTop: 8 }}>
                    <p style={{ fontSize: '0.85em', color: '#475569', margin: '0 0 4px' }}>生徒を選択（複数可）:</p>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, maxHeight: 120, overflowY: 'auto', border: '1px solid #e2e8f0', borderRadius: 4, padding: 6 }}>
                      {masterData.students.map((s) => (
                        <label key={s.id} style={{ display: 'flex', alignItems: 'center', gap: 3, fontSize: '0.85em', cursor: 'pointer' }}>
                          <input type="checkbox" checked={groupStudentIds.includes(s.id)}
                            onChange={(e) => setGroupStudentIds((prev) => e.target.checked ? [...prev, s.id] : prev.filter((id) => id !== s.id))} />
                          {s.name}
                        </label>
                      ))}
                    </div>
                  </div>
                  <p className="muted">集団授業は一科目・講師1人・生徒複数で毎週指定曜日の午前枠で実施されます。</p>
                  <table className="table">
                    <thead><tr><th>講師</th><th>科目</th><th>生徒</th><th>曜日</th><th>時限</th><th>操作</th></tr></thead>
                    <tbody>
                      {(masterData.groupLessons ?? []).map((l) => {
                        const dayNames = ['日', '月', '火', '水', '木', '金', '土']
                        const studentNames = l.studentIds.map((sid) => masterData.students.find((s) => s.id === sid)?.name ?? '?').join(', ')
                        return (
                          <tr key={l.id}>
                            <td>{masterData.teachers.find((t) => t.id === l.teacherId)?.name ?? '-'}</td>
                            <td>{l.subject}</td>
                            <td style={{ fontSize: '0.85em', maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis' }}>{studentNames}</td>
                            <td>{dayNames[l.dayOfWeek]}曜</td><td>午前</td>
                            <td>
                              <button className="btn secondary" type="button" style={{ marginRight: '4px' }} onClick={() => startEditGroupLesson(l)}>編集</button>
                              <button className="btn secondary" type="button" onClick={() => void removeGroupLesson(l.id)}>削除</button>
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="panel">
                  <h3>ペア制約（講師×生徒 / 生徒×生徒）</h3>
                  <div className="row" style={{ flexWrap: 'wrap', gap: '8px', alignItems: 'center' }}>
                    <select value={constraintPersonAType} onChange={(e) => { setConstraintPersonAType(e.target.value as PairConstraintPersonType); setConstraintPersonAId('') }}>
                      <option value="teacher">講師</option>
                      <option value="student">生徒</option>
                    </select>
                    <select value={constraintPersonAId} onChange={(e) => setConstraintPersonAId(e.target.value)}>
                      <option value="">人物Aを選択</option>
                      {(constraintPersonAType === 'teacher' ? masterData.teachers : masterData.students).map((p) => (
                        <option key={p.id} value={p.id}>{p.name}</option>
                      ))}
                    </select>
                    <span style={{ fontSize: '16px', fontWeight: 'bold' }}>×</span>
                    <select value={constraintPersonBType} onChange={(e) => { setConstraintPersonBType(e.target.value as PairConstraintPersonType); setConstraintPersonBId('') }}>
                      <option value="teacher">講師</option>
                      <option value="student">生徒</option>
                    </select>
                    <select value={constraintPersonBId} onChange={(e) => setConstraintPersonBId(e.target.value)}>
                      <option value="">人物Bを選択</option>
                      {(constraintPersonBType === 'teacher' ? masterData.teachers : masterData.students).map((p) => (
                        <option key={p.id} value={p.id}>{p.name}</option>
                      ))}
                    </select>
                    <select value={constraintType} onChange={(e) => setConstraintType(e.target.value as ConstraintType)}>
                      <option value="incompatible">組み合わせ不可</option>
                    </select>
                    {editingConstraintId ? (
                      <>
                        <button className="btn" type="button" onClick={() => void saveEditConstraint()}>更新</button>
                        <button className="btn secondary" type="button" onClick={cancelEditConstraint}>キャンセル</button>
                      </>
                    ) : (
                      <button className="btn" type="button" onClick={() => void upsertConstraint()}>保存</button>
                    )}
                  </div>
                  <table className="table">
                    <thead><tr><th>人物A</th><th></th><th>人物B</th><th>種別</th><th>操作</th></tr></thead>
                    <tbody>
                      {masterData.constraints.map((c) => {
                        const personAName = c.personAType === 'teacher'
                          ? masterData.teachers.find((t) => t.id === c.personAId)?.name
                          : masterData.students.find((s) => s.id === c.personAId)?.name
                        const personBName = c.personBType === 'teacher'
                          ? masterData.teachers.find((t) => t.id === c.personBId)?.name
                          : masterData.students.find((s) => s.id === c.personBId)?.name
                        const personALabel = c.personAType === 'teacher' ? '講師' : '生徒'
                        const personBLabel = c.personBType === 'teacher' ? '講師' : '生徒'
                        return (
                          <tr key={c.id}>
                            <td>{personAName ?? '-'}<span className="muted" style={{ fontSize: '11px', marginLeft: '4px' }}>({personALabel})</span></td>
                            <td style={{ textAlign: 'center' }}>×</td>
                            <td>{personBName ?? '-'}<span className="muted" style={{ fontSize: '11px', marginLeft: '4px' }}>({personBLabel})</span></td>
                            <td><span className="badge warn">不可</span></td>
                            <td>
                              <button className="btn secondary" type="button" style={{ marginRight: '4px' }} onClick={() => startEditConstraint(c)}>編集</button>
                              <button className="btn secondary" type="button" onClick={() => void removeConstraint(c.id)}>削除</button>
                            </td>
                          </tr>
                        )
                      })}
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

  // Precompute helpers (use effective assignments: actual results override planned)
  const effectiveAssignmentsMap = useMemo(
    () => buildEffectiveAssignments(data.assignments, data.actualResults),
    [data.assignments, data.actualResults],
  )
  const allSlotAssignments = useMemo(() => {
    const entries: { slot: string; assignment: Assignment }[] = []
    for (const slot of slotKeys) {
      for (const a of effectiveAssignmentsMap[slot] ?? []) {
        entries.push({ slot, assignment: a })
      }
    }
    return entries
  }, [effectiveAssignmentsMap, slotKeys])

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
    const actualResults = data.actualResults ?? {}
    return data.students.filter((s) => s.submittedAt > 0).map((student) => {
      const myAssignments = allSlotAssignments.filter((e) =>
        e.assignment.studentIds.includes(student.id),
      )
      const totalSlots = myAssignments.filter((e) => !e.assignment.isRegular && !e.assignment.regularMakeupInfo?.[student.id]).length
      const regularSlots = myAssignments.filter((e) => e.assignment.isRegular).length
      const dates = new Set(myAssignments.map((e) => e.slot.split('_')[0]))

      // Count unsatisfied slots: student was in planned assignment but removed from actual result
      // Includes both regular and special lesson students
      let unsatisfiedSlots = 0
      for (const slot of slotKeys) {
        if (!(slot in actualResults)) continue
        const planned = data.assignments[slot] ?? []
        const wasPlanned = planned.some((a) => a.studentIds.includes(student.id))
        // Fallback: check regularLessons for corrupted data
        const [slotDate, slotIdx] = slot.split('_')
        const skDow = getIsoDayOfWeek(slotDate)
        const skSlotNum = Number(slotIdx)
        const wasRegularLesson = data.regularLessons.some((rl) =>
          rl.studentIds.includes(student.id) && rl.dayOfWeek === skDow && rl.slotNumber === skSlotNum
        )
        if (!wasPlanned && !wasRegularLesson) continue
        const actual = actualResults[slot]
        const isInActual = actual.some((r) => r.studentIds.includes(student.id))
        if (!isInActual) unsatisfiedSlots++
      }
      // Add unplaced makeup from autoAssignHighlights (for absences not detectable from corrupted data)
      const unplacedForStudent = (data.autoAssignHighlights?.unplacedMakeup ?? []).filter((um) => um.studentId === student.id)
      if (unplacedForStudent.length > unsatisfiedSlots) {
        unsatisfiedSlots = unplacedForStudent.length
      }

      // Per-subject desired vs assigned (using per-student subjects, excluding makeup)
      const subjectDetails = Object.entries(student.subjectSlots).map(([subj, desired]) => {
        const assigned = myAssignments.filter((e) => !e.assignment.isRegular && !e.assignment.regularMakeupInfo?.[student.id] && getStudentSubject(e.assignment, student.id) === subj).length
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
        unsatisfiedSlots,
      }
    })
  }, [data, allSlotAssignments, slotKeys])

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
      // Count per-student subjects (exclude makeup assignments)
      for (const sid of e.assignment.studentIds) {
        if (e.assignment.regularMakeupInfo?.[sid]) continue
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
                <th>不充足</th>
                <th>科目別（割当/希望）</th>
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
                  <td>
                    {ss.unsatisfiedSlots > 0 ? (
                      <span style={{ color: '#dc2626', fontWeight: 600 }}>{ss.unsatisfiedSlots}コマ</span>
                    ) : (
                      <span style={{ color: '#6b7280' }}>-</span>
                    )}
                  </td>
                  <td style={{ fontSize: '0.85em' }}>
                    {ss.subjectDetails.map((sd) => {
                      const color = sd.diff > 0 ? '#dc2626' : sd.diff < 0 ? '#d97706' : '#16a34a'
                      return (
                        <span key={sd.subject} style={{ marginRight: '8px' }}>
                          {sd.subject}:<span style={{ color, fontWeight: 500 }}>{sd.assigned}/{sd.desired}</span>
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
  const { classroomId = '', sessionId = 'main' } = useParams()
  const navigate = useNavigate()
  const location = useLocation()
  const skipAuth = (location.state as { skipAuth?: boolean } | null)?.skipAuth === true
  const { data, setData, loading, error: sessionError } = useSessionData(classroomId, sessionId)
  const [authorized, setAuthorized] = useState(import.meta.env.DEV || skipAuth)
  const [recentlyUpdated, setRecentlyUpdated] = useState<Set<string>>(new Set())
  const [dragInfo, setDragInfo] = useState<{ sourceSlot: string; sourceIdx: number; teacherId: string; studentIds: string[]; studentDragId?: string; studentDragSubject?: string; regularTeacherId?: string } | null>(null)
  const [, setTransferSlot] = useState<string | null>(null)
  const [showAnalytics, setShowAnalytics] = useState(false)
  const [showRules, setShowRules] = useState(false)
  const [emailSendLog, setEmailSendLog] = useState<Record<string, { time: string; type: string }>>({})
  // --- Actual result recording ---
  const [recordingSlot, setRecordingSlot] = useState<string | null>(null)
  const [editingResults, setEditingResults] = useState<(ActualResult & { _uid?: number })[]>([])
  // --- Salary calculation ---
  const [showSalary, setShowSalary] = useState(false)
  const [showSubmissionPicker, setShowSubmissionPicker] = useState(false)
  const [autoAssignLoading, setAutoAssignLoading] = useState(false)
  const [autoAssignProgress, setAutoAssignProgress] = useState(0)
  const [statusModal, setStatusModal] = useState<StatusReport | null>(null)
  const [latestStatusReport, setLatestStatusReport] = useState<StatusReport | null>(null)
  // Track slots manually modified by user (student move, pair delete, etc.) so auto-fill doesn't overwrite
  const [manuallyModifiedSlots, setManuallyModifiedSlots] = useState<Set<string>>(new Set())
  // --- Slot constraint editing ---
  const [constraintEditStudentId, setConstraintEditStudentId] = useState<string | null>(null)
  const prevSnapshotRef = useRef<{ availability: Record<string, string[]>; studentSubmittedAt: Record<string, number> } | null>(null)
  const masterSyncDoneRef = useRef(false)

  // Sync master data (regularLessons, constraints) into session on first load
  useEffect(() => {
    if (!data || masterSyncDoneRef.current) return
    masterSyncDoneRef.current = true
    const isMendanSession = data.settings.sessionType === 'mendan'
    loadMasterData(classroomId).then((master) => {
      if (!master) return
      const needsUpdate =
        JSON.stringify(data.constraints) !== JSON.stringify(master.constraints) ||
        JSON.stringify(data.regularLessons) !== JSON.stringify(isMendanSession ? [] : master.regularLessons)
      if (!needsUpdate) return
      const next: SessionData = {
        ...data,
        constraints: master.constraints,
        regularLessons: isMendanSession ? [] : master.regularLessons,
      }
      setData(next)
      saveSession(classroomId, sessionId, next).catch(() => { /* ignore */ })
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
    return `${base}/#/c/${classroomId}/availability/${sessionId}/${personType}/${personId}`
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
    navigate(`/c/${classroomId}/availability/${sessionId}/${personType}/${personId}`, { state: { fromAdminInput: true } })
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
    await saveSession(classroomId, sessionId, next)
  }

  const update = async (updater: (current: SessionData) => SessionData): Promise<void> => {
    if (!data) return
    await persist(updater(data))
  }

  // ── Bulk random fill (DEV) ──
  const bulkRandomInstructors = async () => {
    if (!data) return
    const dates = getDatesInRange(data.settings)
    if (dates.length === 0) { alert('講習期間が未設定です'); return }
    const nextAvailability = { ...data.availability }
    const nextSubmittedAt = { ...(data.teacherSubmittedAt ?? {}) }
    for (const inst of instructors) {
      const prefix = isMendan ? 'manager' : 'teacher'
      const key = personKey(prefix as 'teacher' | 'manager', inst.id)
      // Compute regular lesson slots for this teacher
      const regularKeys = new Set<string>()
      if (!isMendan) {
        const teacherLessons = data.regularLessons.filter((l) => l.teacherId === inst.id)
        for (const date of dates) {
          const dow = getIsoDayOfWeek(date)
          for (const lesson of teacherLessons) {
            if (lesson.dayOfWeek === dow) regularKeys.add(`${date}_${lesson.slotNumber}`)
          }
        }
      }
      const avail: string[] = [...regularKeys]
      for (const date of dates) {
        for (let s = 1; s <= data.settings.slotsPerDay; s++) {
          const sk = `${date}_${s}`
          if (!regularKeys.has(sk) && Math.random() < 0.3) avail.push(sk)
        }
      }
      nextAvailability[key] = avail
      nextSubmittedAt[inst.id] = Date.now()
    }
    await persist({ ...data, availability: nextAvailability, teacherSubmittedAt: nextSubmittedAt })
  }

  const bulkRandomStudents = async () => {
    if (!data) return
    const dates = getDatesInRange(data.settings)
    if (dates.length === 0) { alert('講習期間が未設定です'); return }
    if (isMendan) {
      // Mendan: random parent availability from manager-available slots
      const nextAvailability = { ...data.availability }
      const updatedStudents = data.students.map((student) => {
        const key = personKey('student', student.id)
        // Pick ~50% of available slots
        const avail: string[] = []
        for (const date of dates) {
          for (let s = 1; s <= data.settings.slotsPerDay; s++) {
            if (Math.random() < 0.5) avail.push(`${date}_${s}`)
          }
        }
        nextAvailability[key] = avail
        return { ...student, submittedAt: Date.now() }
      })
      await persist({ ...data, students: updatedStudents, availability: nextAvailability })
    } else {
      // Normal: random subjects + unavailable slots
      const SUBJ = ['英', '数', '国', '理', '社', 'IT', '算']
      const updatedStudents = data.students.map((student) => {
        const shuffled = [...SUBJ].sort(() => Math.random() - 0.5)
        const count = 1 + Math.floor(Math.random() * 3)
        const subjectSlots: Record<string, number> = {}
        const subjects: string[] = []
        for (let i = 0; i < count && i < shuffled.length; i++) {
          const slots = 1 + Math.floor(Math.random() * 4)
          subjectSlots[shuffled[i]] = slots
          subjects.push(shuffled[i])
        }
        // Random ~20% unavailable slots
        const unavailable: string[] = []
        for (const date of dates) {
          for (let s = 1; s <= data.settings.slotsPerDay; s++) {
            if (Math.random() < 0.2) unavailable.push(`${date}_${s}`)
          }
        }
        // Derive unavailable dates
        const dateSlotCounts = new Map<string, number>()
        for (const sk of unavailable) {
          const d = sk.split('_')[0]
          dateSlotCounts.set(d, (dateSlotCounts.get(d) ?? 0) + 1)
        }
        const unavailDates = [...dateSlotCounts.entries()]
          .filter(([, c]) => c >= data.settings.slotsPerDay)
          .map(([d]) => d)
        return {
          ...student,
          subjects,
          subjectSlots,
          unavailableDates: unavailDates,
          preferredSlots: [],
          unavailableSlots: unavailable,
          submittedAt: Date.now(),
        }
      })
      await persist({ ...data, students: updatedStudents })
    }
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
      const verified = await saveAndVerify(classroomId, sessionId, seed)
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
      const master = await loadMasterData(classroomId)
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
        regularLessons: master.regularLessons,
        groupLessons: master.groupLessons ?? [],
      }
      // Only save if something actually changed
      const changed =
        JSON.stringify(data.managers ?? []) !== JSON.stringify(next.managers) ||
        JSON.stringify(data.teachers) !== JSON.stringify(next.teachers) ||
        JSON.stringify(data.students.map((s) => ({ id: s.id, name: s.name, grade: s.grade, memo: s.memo }))) !==
          JSON.stringify(next.students.map((s) => ({ id: s.id, name: s.name, grade: s.grade, memo: s.memo }))) ||
        JSON.stringify(data.constraints) !== JSON.stringify(next.constraints) ||
        JSON.stringify(data.regularLessons) !== JSON.stringify(next.regularLessons) ||
        JSON.stringify(data.groupLessons ?? []) !== JSON.stringify(next.groupLessons)
      if (changed) {
        setData(next)
        await saveSession(classroomId, sessionId, next)
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
    // Include student unavailability in signature so auto-fill reacts to absent students
    const studentUnavailSig = data.students.map(s => `${s.id}:${(s.unavailableSlots ?? []).join(';')}`).join('|')
    const sig = `${slotKeys.join(',')}|${data.regularLessons.map((l) => `${l.id}:${l.dayOfWeek}:${l.slotNumber}:${l.teacherId}:${l.studentIds.join('+')}:${l.subject}:${JSON.stringify(l.studentSubjects ?? {})}`).join(',')}|${(data.groupLessons ?? []).map((l) => `G${l.id}:${l.dayOfWeek}:${l.slotNumber}:${l.teacherId}:${l.studentIds.join('+')}:${l.subject}`).join(',')}|${studentUnavailSig}|${assignmentStateSig}`
    if (sig === regularFillSigRef.current) return
    regularFillSigRef.current = sig

    let changed = false
    const nextAssignments = { ...data.assignments }

    for (const slot of slotKeys) {
      const existing = nextAssignments[slot]
      const slotRegularLessons = findRegularLessonsForSlot(data.regularLessons, slot)

      if (slotRegularLessons.length === 0) {
        if (existing && existing.length > 0 && existing.every((a) => a.isRegular && !a.isGroupLesson)) {
          delete nextAssignments[slot]
          changed = true
        }
        continue
      }

      // Don't overwrite manual (non-regular) assignments or manually modified slots
      if (manuallyModifiedSlots.has(slot)) continue
      if (existing && existing.some((a) => hasMeaningfulManualAssignment(a))) continue
      // Don't overwrite slots where a student was manually transferred out (regularMakeupInfo present)
      if (existing && existing.some((a) => a.regularMakeupInfo && Object.keys(a.regularMakeupInfo).length > 0)) continue

      const expectedRegulars = slotRegularLessons.map((lesson) => {
        // Exclude students who are unavailable for this slot (absent)
        const availableStudentIds = lesson.studentIds.filter((sid) => {
          const student = data.students.find((s) => s.id === sid)
          if (!student) return false
          return isStudentAvailable(student, slot)
        })
        return {
          teacherId: lesson.teacherId,
          studentIds: availableStudentIds,
          subject: lesson.studentSubjects?.[lesson.studentIds[0]] ?? lesson.subject,
          studentSubjects: lesson.studentSubjects,
          isRegular: true,
        }
      }).filter((a) => a.studentIds.length > 0) // Skip teacher-only assignments when all students are unavailable

      const expectedSig = expectedRegulars
        .map((a) => assignmentSignature(a))
        .sort()
      const existingSig = (existing ?? [])
        .filter((a) => a.isRegular && !a.isGroupLesson)
        .map((a) => assignmentSignature(a))
        .sort()

      if (expectedSig.length === existingSig.length && expectedSig.every((sigItem, idx) => sigItem === existingSig[idx])) {
        continue
      }

      // Don't overwrite if existing regular assignments are a subset of expected
      // (students may have been manually transferred out via drag)
      const existingRegulars = (existing ?? []).filter((a) => a.isRegular && !a.isGroupLesson)
      const existingRegularStudentIds = new Set(existingRegulars.flatMap((a) => a.studentIds))
      const expectedStudentIds = new Set(expectedRegulars.flatMap((a) => a.studentIds))
      if (existingRegulars.length > 0 && [...existingRegularStudentIds].every((sid) => expectedStudentIds.has(sid)) && existingRegularStudentIds.size < expectedStudentIds.size) {
        continue
      }

      // Preserve any non-regular assignments already in the slot (e.g. special students added)
      const nonRegularExisting = (existing ?? []).filter((a) => !a.isRegular || a.isGroupLesson)
      nextAssignments[slot] = [...nonRegularExisting, ...expectedRegulars]
      changed = true
    }

    // Auto-fill group lessons for matching day-of-week and slot number
    for (const slot of slotKeys) {
      const dayOfWeek = getSlotDayOfWeek(slot)
      const slotNumber = getSlotNumber(slot)
      const matchingGroupLessons = (data.groupLessons ?? []).filter((gl) => gl.dayOfWeek === dayOfWeek && gl.slotNumber === slotNumber)
      if (matchingGroupLessons.length === 0) continue

      const existing = nextAssignments[slot] ?? []
      for (const gl of matchingGroupLessons) {
        // Check if this group lesson is already present
        const alreadyPresent = existing.some((a) => a.isGroupLesson && a.teacherId === gl.teacherId && a.studentIds.length === gl.studentIds.length && gl.studentIds.every((sid) => a.studentIds.includes(sid)))
        if (alreadyPresent) continue

        existing.push({
          teacherId: gl.teacherId,
          studentIds: [...gl.studentIds],
          subject: gl.subject,
          isRegular: true,
          isGroupLesson: true,
        })
        changed = true
      }
      nextAssignments[slot] = existing
    }

    // Auto-register regular lesson and group lesson slots as teacher availability
    const nextAvailability = { ...data.availability }
    let availChanged = false
    const allAutoLessons = [...data.regularLessons, ...(data.groupLessons ?? []).map((gl) => ({ ...gl, studentSubjects: undefined }))]
    for (const lesson of allAutoLessons) {
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

  // --- Actual result recording helpers ---
  let editingResultUid = 0
  const startRecording = (slot: string): void => {
    if (!data) return
    const assignments = data.assignments[slot] ?? []
    const existing = data.actualResults?.[slot]
    if (existing) {
      // Edit mode: load existing results
      setEditingResults(existing.map((r) => ({ ...r, studentIds: [...r.studentIds], _uid: ++editingResultUid })))
    } else {
      // New: copy from current assignments
      setEditingResults(assignments.map((a) => ({
        teacherId: a.teacherId,
        studentIds: [...a.studentIds],
        subject: a.subject,
        ...(a.studentSubjects ? { studentSubjects: { ...a.studentSubjects } } : {}),
        _uid: ++editingResultUid,
      })))
    }
    setRecordingSlot(slot)
  }

  const saveActualResults = async (): Promise<void> => {
    if (!data || !recordingSlot) return
    // Strip _uid and undefined fields before persisting (Firestore rejects undefined)
    const cleaned: ActualResult[] = editingResults.map(({ _uid: _, ...rest }) => {
      const result: ActualResult = { teacherId: rest.teacherId, studentIds: rest.studentIds, subject: rest.subject }
      if (rest.studentSubjects && Object.keys(rest.studentSubjects).length > 0) {
        result.studentSubjects = rest.studentSubjects
      }
      return result
    })
    const nextResults = { ...(data.actualResults ?? {}), [recordingSlot]: cleaned }
    setRecordingSlot(null)
    setEditingResults([])
    await persist({ ...data, actualResults: nextResults })
  }

  const cancelRecording = (): void => {
    setRecordingSlot(null)
    setEditingResults([])
  }

  const clearActualResults = async (slot: string): Promise<void> => {
    if (!data) return
    if (!window.confirm('この実績記録を解除しますか？\n割当は元の計画に戻ります。')) return
    const nextResults = { ...(data.actualResults ?? {}) }
    delete nextResults[slot]
    // Use empty object instead of undefined (Firestore rejects undefined values)
    await persist({ ...data, actualResults: Object.keys(nextResults).length > 0 ? nextResults : {} })
  }

  const updateEditingResult = (idx: number, field: keyof ActualResult, value: string | string[]): void => {
    setEditingResults((prev) => prev.map((r, i) => {
      if (i !== idx) return r
      if (field === 'teacherId') return { ...r, teacherId: value as string }
      if (field === 'studentIds') return { ...r, studentIds: value as string[] }
      if (field === 'subject') return { ...r, subject: value as string }
      return r
    }))
  }

  const updateEditingResultStudentSubject = (idx: number, studentId: string, subject: string): void => {
    setEditingResults((prev) => prev.map((r, i) => {
      if (i !== idx) return r
      const ss = { ...(r.studentSubjects ?? {}) }
      ss[studentId] = subject
      return { ...r, studentSubjects: ss }
    }))
  }

  const addEditingResultPair = (): void => {
    setEditingResults((prev) => [...prev, { teacherId: '', studentIds: [''], subject: '', studentSubjects: {}, _uid: ++editingResultUid }])
  }

  const removeEditingResultPair = (idx: number): void => {
    setEditingResults((prev) => prev.filter((_, i) => i !== idx))
  }

  // --- Salary calculation helpers ---
  type SalaryTier = 'A' | 'B' | 'C' | 'D'
  const defaultTierRates = { A: 0, B: 0, C: 0, D: 0 }

  const isHighSchoolOrAbove = (grade: string): boolean => grade.startsWith('高')

  const classifyTier = (result: ActualResult): SalaryTier => {
    if (!data) return 'A'
    const studentCount = result.studentIds.filter((s) => s).length
    const hasHighSchool = result.studentIds.some((sid) => {
      const student = data.students.find((s) => s.id === sid)
      return student ? isHighSchoolOrAbove(student.grade) : false
    })
    if (studentCount <= 1 && !hasHighSchool) return 'A'
    if (studentCount >= 2 && !hasHighSchool) return 'B'
    if (studentCount <= 1 && hasHighSchool) return 'C'
    return 'D' // 2 students, at least one high school
  }

  const saveTierRate = async (tier: SalaryTier, rate: number): Promise<void> => {
    if (!data) return
    const current = data.tierRates ?? { ...defaultTierRates }
    await persist({ ...data, tierRates: { ...current, [tier]: rate } })
  }

  type SalaryRow = { teacherId: string; name: string; A: number; B: number; C: number; D: number; total: number }
  const computeSalaryData = (rates: { A: number; B: number; C: number; D: number }): SalaryRow[] => {
    if (!data) return []
    const results = data.actualResults ?? {}
    const assignments = data.assignments ?? {}
    const tierCountMap: Record<string, { A: number; B: number; C: number; D: number }> = {}
    for (const slot of Object.keys(results)) {
      for (const r of results[slot]) {
        if (!r.teacherId) continue
        // Skip regular lesson slots (check original assignments)
        const origAssign = (assignments[slot] ?? []).find((a) => a.teacherId === r.teacherId)
        if (origAssign?.isRegular) continue
        if (!tierCountMap[r.teacherId]) tierCountMap[r.teacherId] = { A: 0, B: 0, C: 0, D: 0 }
        const tier = classifyTier(r)
        tierCountMap[r.teacherId][tier]++
      }
    }
    return instructors
      .filter((t) => tierCountMap[t.id])
      .map((t) => {
        const counts = tierCountMap[t.id]
        const total = counts.A * rates.A + counts.B * rates.B + counts.C * rates.C + counts.D * rates.D
        return { teacherId: t.id, name: t.name, ...counts, total }
      })
      .sort((a, b) => (b.A + b.B + b.C + b.D) - (a.A + a.B + a.C + a.D))
  }

  const buildConstraintSuggestion = (student: Student, blockReason: string, slot: string, teacherName: string): StatusProposal | null => {
    const cards = student.constraintCards ?? getDefaultConstraintCards(student.grade)
    if (blockReason.startsWith('一コマ限定') && cards.includes('oneSlotOnly')) {
      return toStatusProposal(`制約カード変更案: 一コマ限定 -> 二コマ限定 または 三コマ限定 に変更すると ${slotLabel(slot, isMendan, mendanStart)} で講師(${teacherName})に追加候補`)
    }
    if (blockReason.startsWith('二コマ限定') && cards.includes('twoSlotLimit')) {
      return toStatusProposal(`制約カード変更案: 二コマ限定 -> 三コマ限定 に変更すると ${slotLabel(slot, isMendan, mendanStart)} で講師(${teacherName})に追加候補`)
    }
    if (blockReason.startsWith('2コマ連続') && cards.includes('twoConsecutive')) {
      return toStatusProposal(`制約カード変更案: 2コマ連続 を外すか 二コマ限定 に変更すると ${slotLabel(slot, isMendan, mendanStart)} で講師(${teacherName})に追加候補`)
    }
    if (blockReason.startsWith('2コマ連続(一コマ空け)') && cards.includes('twoWithGap')) {
      return toStatusProposal(`制約カード変更案: 2コマ連続(一コマ空け) を外すか 二コマ限定 に変更すると ${slotLabel(slot, isMendan, mendanStart)} で講師(${teacherName})に追加候補`)
    }
    if (blockReason.startsWith('通常授業連結') && cards.includes('regularLink')) {
      return toStatusProposal(`制約カード変更案: 通常授業連結 を外すか 二コマ限定 に変更すると ${slotLabel(slot, isMendan, mendanStart)} で講師(${teacherName})に追加候補`)
    }
    if (blockReason.startsWith('集団後連続') && cards.includes('groupContinuous')) {
      return toStatusProposal(`制約カード変更案: 集団後連続 を外すと ${slotLabel(slot, isMendan, mendanStart)} で講師(${teacherName})に追加候補`)
    }
    return null
  }

  const collectPendingMakeupDemands = (
    assignmentState: Record<string, Assignment[]>,
    actualResultsOverride?: Record<string, ActualResult[]>,
  ): PendingMakeupDemand[] => {
    if (!data || isMendan) return []

    const actualResults = actualResultsOverride ?? data.actualResults
    const allSlotKeys = [...new Set([...slotKeys, ...Object.keys(actualResults ?? {})])]
    const demands: PendingMakeupDemand[] = []

    for (const rl of data.regularLessons) {
      for (const sid of rl.studentIds) {
        const student = data.students.find((s) => s.id === sid)
        if (!student) continue
        const rlSubject = rl.studentSubjects?.[sid] ?? rl.subject
        for (const slot of allSlotKeys) {
          const [date] = slot.split('_')
          const dow = getIsoDayOfWeek(date)
          if (dow !== rl.dayOfWeek || getSlotNumber(slot) !== rl.slotNumber) continue
          const actualForSlot = actualResults?.[slot]
          const absentFromActual = actualForSlot != null && !actualForSlot.some((r) => r.studentIds.includes(sid))
          const needsMakeup = !isStudentAvailable(student, slot)
          if (!needsMakeup && !absentFromActual) continue
          demands.push({
            studentId: sid,
            teacherId: rl.teacherId,
            subject: rlSubject,
            ...(absentFromActual ? { absentDate: date } : {}),
            makeupInfo: { dayOfWeek: rl.dayOfWeek, slotNumber: rl.slotNumber, date },
          })
        }
      }
    }

    if (actualResults) {
      for (const [slot, results] of Object.entries(actualResults)) {
        const origAssignments = assignmentState[slot] ?? data.assignments[slot] ?? []
        for (const orig of origAssignments) {
          if (orig.isRegular || !orig.regularMakeupInfo) continue
          for (const sid of orig.studentIds) {
            const makeupInfo = orig.regularMakeupInfo[sid]
            if (!makeupInfo) continue
            if (results.some((r) => r.studentIds.includes(sid))) continue
            demands.push({
              studentId: sid,
              teacherId: orig.teacherId,
              subject: getStudentSubject(orig, sid),
              absentDate: slot.split('_')[0],
              makeupInfo,
            })
          }
        }
      }
    }

    const pending = [...demands]
    for (const slotAssignments of Object.values(assignmentState)) {
      for (const assignment of slotAssignments) {
        if (!assignment.regularMakeupInfo) continue
        for (const [sid, info] of Object.entries(assignment.regularMakeupInfo)) {
          const subject = getStudentSubject(assignment, sid)
          const matchedIndex = pending.findIndex((demand) =>
            demand.studentId === sid
            && demand.subject === subject
            && demand.makeupInfo.dayOfWeek === info.dayOfWeek
            && demand.makeupInfo.slotNumber === info.slotNumber,
          )
          if (matchedIndex >= 0) pending.splice(matchedIndex, 1)
        }
      }
    }

    return pending
  }

  const analyzePlacementOptions = (
    assignmentState: Record<string, Assignment[]>,
    effectiveAssignments: Record<string, Assignment[]>,
    availableSlots: string[],
    student: Student,
    subject: string,
    candidateTeachers: Teacher[],
    options?: { slotFilter?: (slot: string) => boolean; makeupDemand?: PendingMakeupDemand },
  ): PlacementAnalysis => {
    if (!data) return { force: [], teacher: [], student: [], cards: [], blockers: [] }

    const forceSuggestions = new Map<string, StatusProposal>()
    const teacherSuggestions = new Set<string>()
    const studentSuggestions = new Set<string>()
    const cardSuggestions = new Map<string, StatusProposal>()
    const blockerReasons = new Set<string>()
    const deskLimit = data.settings.deskCount ?? 0

    for (const slot of availableSlots) {
      if (getSlotNumber(slot) === 0) continue
      if (options?.slotFilter && !options.slotFilter(slot)) continue
      const slotAssignments = assignmentState[slot] ?? []
      const [slotDate] = slot.split('_')
      const dayOfWeek = getIsoDayOfWeek(slotDate)
      const groupLessonsOnDate = (data.groupLessons ?? []).filter((gl) => gl.dayOfWeek === dayOfWeek)

      for (const teacher of candidateTeachers) {
        if (slotAssignments.some((a) => a.studentIds.includes(student.id))) continue

        const teacherAssignment = slotAssignments.find((a) => a.teacherId === teacher.id && !a.isGroupLesson)
        const teacherAvailable = hasAvailability(data.availability, instructorPersonType, teacher.id, slot)
        const studentAvailable = isStudentAvailable(student, slot)
        const teacherStudentIncompatible = constraintFor(data.constraints, teacher.id, student.id) === 'incompatible'
        const existingStudentIncompatible = !!teacherAssignment?.studentIds.length && teacherAssignment.studentIds.some((sid) => constraintFor(data.constraints, sid, student.id) === 'incompatible')
        const teacherPairFull = !!teacherAssignment && teacherAssignment.studentIds.length >= 2
        const deskBlocked = !teacherAssignment && deskLimit > 0 && slotAssignments.length >= deskLimit
        const evalResult = evaluateConstraintCards(
          student,
          slot,
          effectiveAssignments,
          data.settings.slotsPerDay,
          data.regularLessons,
          groupLessonsOnDate,
          teacher.id,
        )

        const hardReasons: string[] = []
        if (teacherStudentIncompatible) hardReasons.push(`講師(${teacher.name})と生徒の相性不可`)
        if (existingStudentIncompatible) hardReasons.push(`講師(${teacher.name})ペア内の既存生徒と相性不可`)
        if (teacherPairFull) hardReasons.push(`講師(${teacher.name})のペアが満席`)
        if (deskBlocked) hardReasons.push('机数上限で新規ペア不可')
        if (evalResult.blocked && evalResult.blockReason) hardReasons.push(evalResult.blockReason)

        if (teacherAvailable && studentAvailable && hardReasons.length > 0) {
          for (const reason of hardReasons) blockerReasons.add(`${slotLabel(slot, isMendan, mendanStart)}: ${reason}`)
        }

        if (teacherAvailable && studentAvailable && !teacherStudentIncompatible && !existingStudentIncompatible && !teacherPairFull && !deskBlocked && evalResult.blocked) {
          const verb = options?.makeupDemand ? '強制振替' : '強制割当'
          const targetText = teacherAssignment ? '既存ペアへ追加' : '新規ペアで追加'
          const label = `制約違反${verb}案: ${slotLabel(slot, isMendan, mendanStart)} / ${teacher.name} / ${subject} / ${targetText}${evalResult.blockReason ? ` (${evalResult.blockReason})` : ''}`
          const action: ForceAssignAction = {
            type: 'force-assign',
            slot,
            teacherId: teacher.id,
            studentId: student.id,
            subject,
            ...(options?.makeupDemand ? { makeupInfo: options.makeupDemand.makeupInfo } : {}),
          }
          forceSuggestions.set(`${slot}|${teacher.id}|${student.id}|${subject}|${options?.makeupDemand ? 'makeup' : 'normal'}`, toStatusProposal(label, action))
        }

        if (!teacherAvailable && studentAvailable && !teacherStudentIncompatible && !existingStudentIncompatible && !teacherPairFull && !deskBlocked) {
          teacherSuggestions.add(`講師出席追加案: ${teacher.name} を ${slotLabel(slot, isMendan, mendanStart)} 出席可にすると ${subject} を追加候補`)
        }
        if (teacherAvailable && !studentAvailable && !teacherStudentIncompatible && !existingStudentIncompatible && !teacherPairFull && !deskBlocked) {
          studentSuggestions.add(`生徒出席緩和案: ${slotLabel(slot, isMendan, mendanStart)} を出席可にすると 講師(${teacher.name}) で ${subject} を追加候補`)
        }
        if (teacherAvailable && studentAvailable && evalResult.blocked) {
          const suggestion = buildConstraintSuggestion(student, evalResult.blockReason ?? '', slot, teacher.name)
          if (suggestion) cardSuggestions.set(suggestion.label, suggestion)
        }
        if (!teacherAvailable && !studentAvailable && !teacherStudentIncompatible && !existingStudentIncompatible && !teacherPairFull && !deskBlocked) {
          blockerReasons.add(`${slotLabel(slot, isMendan, mendanStart)}: 講師(${teacher.name})と生徒の両方が未出席`)
        } else if (!teacherAvailable && !teacherStudentIncompatible && !existingStudentIncompatible && !teacherPairFull && !deskBlocked) {
          blockerReasons.add(`${slotLabel(slot, isMendan, mendanStart)}: 講師(${teacher.name})が未出席`)
        } else if (!studentAvailable && !teacherStudentIncompatible && !existingStudentIncompatible && !teacherPairFull && !deskBlocked) {
          blockerReasons.add(`${slotLabel(slot, isMendan, mendanStart)}: 生徒が出席不可`)
        }
      }
    }

    return {
      force: [...forceSuggestions.values()].slice(0, 5),
      teacher: [...teacherSuggestions].slice(0, 5).map((label) => toStatusProposal(label)),
      student: [...studentSuggestions].slice(0, 5).map((label) => toStatusProposal(label)),
      cards: [...cardSuggestions.values()].slice(0, 5),
      blockers: [...blockerReasons].slice(0, 8),
    }
  }

  const buildRemainingSuggestions = (
    assignmentState: Record<string, Assignment[]>,
    effectiveAssignments: Record<string, Assignment[]>,
    availableSlots: string[],
    student: Student,
    subject: string,
  ): PlacementAnalysis => {
    const compatibleTeachers = data?.teachers.filter((teacher) =>
      canTeachSubject(teacher.subjects, student.grade, subject)
      && constraintFor(data.constraints, teacher.id, student.id) !== 'incompatible',
    ) ?? []
    return analyzePlacementOptions(assignmentState, effectiveAssignments, availableSlots, student, subject, compatibleTeachers)
  }

  const applyForceAssignAction = async (proposal: ForceAssignAction): Promise<void> => {
    if (!data) return
    let success = false
    let errorMessage = ''
    const studentName = data.students.find((s) => s.id === proposal.studentId)?.name ?? proposal.studentId
    const teacherName = data.teachers.find((t) => t.id === proposal.teacherId)?.name ?? proposal.teacherId

    await updateAssignments((current) => {
      const student = current.students.find((s) => s.id === proposal.studentId)
      const teacher = current.teachers.find((t) => t.id === proposal.teacherId)
      if (!student || !teacher) {
        errorMessage = '生徒または講師が見つからないため、強制割当できませんでした。'
        return current
      }
      if (!hasAvailability(current.availability, 'teacher', proposal.teacherId, proposal.slot)) {
        errorMessage = `${teacher.name} は現在 ${slotLabel(proposal.slot, isMendan, mendanStart)} に出席不可です。`
        return current
      }
      if (!isStudentAvailable(student, proposal.slot)) {
        errorMessage = `${student.name} は現在 ${slotLabel(proposal.slot, isMendan, mendanStart)} に出席不可です。`
        return current
      }

      const slotAssignments = [...(current.assignments[proposal.slot] ?? [])]
      if (slotAssignments.some((assignment) => assignment.studentIds.includes(proposal.studentId))) {
        errorMessage = `${student.name} は既に ${slotLabel(proposal.slot, isMendan, mendanStart)} に割当済みです。`
        return current
      }

      const targetIndex = slotAssignments.findIndex((assignment) => assignment.teacherId === proposal.teacherId && !assignment.isGroupLesson)
      if (targetIndex >= 0) {
        const targetAssignment = slotAssignments[targetIndex]
        if (targetAssignment.studentIds.length >= 2) {
          errorMessage = `${teacher.name} のペアは既に満席です。`
          return current
        }
        if (targetAssignment.studentIds.some((sid) => constraintFor(current.constraints, sid, proposal.studentId) === 'incompatible')) {
          errorMessage = `${teacher.name} の既存ペアには相性不可の生徒がいるため、強制割当できません。`
          return current
        }
        const nextStudentIds = [...targetAssignment.studentIds, proposal.studentId]
        const nextStudentSubjects = { ...(targetAssignment.studentSubjects ?? {}), [proposal.studentId]: proposal.subject }
        const primarySubject = nextStudentSubjects[nextStudentIds[0]] ?? targetAssignment.subject ?? proposal.subject
        const nextMakeupInfo = proposal.makeupInfo
          ? { ...(targetAssignment.regularMakeupInfo ?? {}), [proposal.studentId]: proposal.makeupInfo }
          : targetAssignment.regularMakeupInfo
        slotAssignments[targetIndex] = {
          ...targetAssignment,
          studentIds: nextStudentIds,
          subject: primarySubject,
          studentSubjects: nextStudentSubjects,
          ...(nextMakeupInfo ? { regularMakeupInfo: nextMakeupInfo } : {}),
        }
      } else {
        const deskCount = current.settings.deskCount ?? 0
        if (deskCount > 0 && slotAssignments.length >= deskCount) {
          errorMessage = '机数上限のため、強制割当できません。'
          return current
        }
        slotAssignments.push({
          teacherId: proposal.teacherId,
          studentIds: [proposal.studentId],
          subject: proposal.subject,
          studentSubjects: { [proposal.studentId]: proposal.subject },
          ...(proposal.makeupInfo ? { regularMakeupInfo: { [proposal.studentId]: proposal.makeupInfo } } : {}),
        })
      }

      success = true
      setManuallyModifiedSlots((prev) => new Set(prev).add(proposal.slot))
      const currentHighlights = current.autoAssignHighlights ?? { added: {}, changed: {}, makeup: {} }
      const filteredUnplacedMakeup = (currentHighlights.unplacedMakeup ?? []).filter((item) => {
        if (item.studentId !== proposal.studentId) return true
        if (item.subject !== proposal.subject) return true
        if (proposal.makeupInfo && item.teacherId !== proposal.teacherId) return true
        return false
      })
      return {
        ...current,
        assignments: { ...current.assignments, [proposal.slot]: slotAssignments },
        autoAssignHighlights: {
          ...currentHighlights,
          ...(currentHighlights.changeDetails ? { changeDetails: currentHighlights.changeDetails } : {}),
          unplacedMakeup: filteredUnplacedMakeup,
        },
      }
    })

    if (errorMessage) {
      alert(errorMessage)
      return
    }
    if (success) {
      setLatestStatusReport(null)
      setStatusModal(null)
      alert(`${proposal.makeupInfo ? '強制振替' : '強制割当'}を実行しました。\n${slotLabel(proposal.slot, isMendan, mendanStart)} / ${teacherName} / ${studentName} / ${proposal.subject}`)
    }
  }

  const applyAutoAssign = async (): Promise<void> => {
    if (!data) return
    setAutoAssignLoading(true)
    setAutoAssignProgress(0)
    setManuallyModifiedSlots(new Set()) // Clear manual modifications on auto-assign
    // Yield to let React render the spinner
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))

    try {

    // Filter out slots with actual results
    const recordedSlots = new Set(Object.keys(data.actualResults ?? {}))
    const availableSlotKeys = slotKeys.filter((s) => !recordedSlots.has(s))
    console.log('[AutoAssign] totalSlots:', slotKeys.length, 'recordedSlots:', recordedSlots.size, 'availableSlots:', availableSlotKeys.length)

    // Mendan FCFS auto-assign
    if (isMendan) {
      const { assignments: nextAssignments, unassignedParents } = buildMendanAutoAssignments(data, availableSlotKeys)
      // Preserve recorded slot assignments from actual results (already in builder result)
      // For recorded slots with empty actual results, set empty array (all absent)
      for (const slot of recordedSlots) {
        if (!nextAssignments[slot]) {
          const origSlot = data.assignments[slot] ?? []
          nextAssignments[slot] = data.actualResults?.[slot]
            ? data.actualResults[slot].map((r) => {
                const orig = origSlot.find((a) => a.teacherId === r.teacherId)
                return {
                  teacherId: r.teacherId,
                  studentIds: [...r.studentIds],
                  subject: r.subject,
                  ...(r.studentSubjects ? { studentSubjects: { ...r.studentSubjects } } : {}),
                  ...(orig?.isRegular ? { isRegular: true } : {}),
                  ...(orig?.isGroupLesson ? { isGroupLesson: true } : {}),
                  ...(orig?.regularMakeupInfo ? { regularMakeupInfo: { ...orig.regularMakeupInfo } } : {}),
                }
              })
            : origSlot
        }
      }
      const submittedCount = data.students.filter((s) => s.submittedAt > 0).length
      const assignedCount = submittedCount - unassignedParents.length

      // Final cleanup: remove teacher-only assignments
      const validStudentIdSet = new Set(data.students.map((s) => s.id))
      for (const slot of Object.keys(nextAssignments)) {
        nextAssignments[slot] = nextAssignments[slot].filter((a) => a.studentIds.some(sid => sid && validStudentIdSet.has(sid)) || a.isGroupLesson)
        if (nextAssignments[slot].length === 0) delete nextAssignments[slot]
      }

      await updateAssignments((current) => ({
        ...current,
        assignments: nextAssignments,
        autoAssignHighlights: { added: {}, changed: {}, makeup: {} },
        settings: { ...current.settings, lastAutoAssignedAt: Date.now() },
      }))

      const msg = unassignedParents.length > 0
        ? `自動割当完了: ${assignedCount}/${submittedCount}名を割当しました。\n\n未割当の保護者:\n${unassignedParents.join('\n')}`
        : `自動割当完了: 提出済み${assignedCount}名全員を割当しました。`
      alert(msg)
      setAutoAssignLoading(false)
      return
    }

    // Lecture auto-assign (existing logic)
    const { assignments: nextAssignments, changeLog, changedPairSignatures, addedPairSignatures, makeupPairSignatures, changeDetails, unplacedMakeup } = await buildIncrementalAutoAssignments(data, availableSlotKeys, (ratio) => setAutoAssignProgress(ratio))

    const highlightAdded: Record<string, string[]> = {}
    const highlightChanged: Record<string, string[]> = {}
    const highlightMakeup: Record<string, string[]> = {}
    const highlightDetails: Record<string, Record<string, string>> = {}
    // Always show 振替 badges for regular student makeup assignments
    for (const slot of Object.keys(makeupPairSignatures)) {
      if ((makeupPairSignatures[slot] ?? []).length > 0) highlightMakeup[slot] = [...makeupPairSignatures[slot]]
      if (changeDetails[slot]) {
        if (!highlightDetails[slot]) highlightDetails[slot] = {}
        Object.assign(highlightDetails[slot], changeDetails[slot])
      }
    }
    // Only show 新規/変更 badges on 2nd+ auto-assign
    const hadPreviousAssignments = !!(data.settings?.lastAutoAssignedAt)
    if (hadPreviousAssignments) {
      const allSlotSet = new Set([...Object.keys(addedPairSignatures), ...Object.keys(changedPairSignatures)])
      for (const slot of allSlotSet) {
        if ((addedPairSignatures[slot] ?? []).length > 0) highlightAdded[slot] = [...addedPairSignatures[slot]]
        if ((changedPairSignatures[slot] ?? []).length > 0) highlightChanged[slot] = [...changedPairSignatures[slot]]
        if (changeDetails[slot]) highlightDetails[slot] = { ...(highlightDetails[slot] ?? {}), ...changeDetails[slot] }
      }
    }

    // Build effective assignments including recorded slots for accurate demand counting
    const effectiveForCounting = buildEffectiveAssignments(
      { ...nextAssignments, ...Object.fromEntries([...recordedSlots].map((s) => [s, data.assignments[s] ?? []])) },
      data.actualResults,
    )

    const remainingStudents = data.students
      .map((student) => {
        const remaining = Object.entries(student.subjectSlots)
          .map(([subject, desired]) => {
            const assigned = countStudentSubjectLoad(effectiveForCounting, student.id, subject)
            return {
              subject,
              remaining: desired - assigned,
              suggestions: buildRemainingSuggestions(nextAssignments, effectiveForCounting, availableSlotKeys, student, subject),
            }
          })
          .filter((item) => item.remaining > 0)
        if (remaining.length === 0) return null
        return { studentName: student.name, remaining }
      })
      .filter(Boolean) as { studentName: string; remaining: { subject: string; remaining: number; suggestions: PlacementAnalysis }[] }[]

    const overAssignedStudents = data.students
      .map((student) => {
        const over = Object.entries(student.subjectSlots)
          .map(([subject, desired]) => {
            const assigned = countStudentSubjectLoad(effectiveForCounting, student.id, subject)
            return { subject, over: assigned - desired }
          })
          .filter((item) => item.over > 0)
        if (over.length === 0) return null
        return { studentName: student.name, over }
      })
      .filter(Boolean) as { studentName: string; over: { subject: string; over: number }[] }[]

    const shortageEntries = collectTeacherShortages(data, nextAssignments)

    // Build unplaced makeup message with reasons
    const pendingMakeupDemands = collectPendingMakeupDemands(nextAssignments)
    const unplacedMakeupEntries: { studentName: string; subject: string; causes: string[]; proposals: StatusProposal[]; reason: string }[] = []
    for (const um of unplacedMakeup) {
      const student = data.students.find((s) => s.id === um.studentId)
      if (!student) continue
      const teacher = data.teachers.find((t) => t.id === um.teacherId)
      const teacherName = teacher?.name ?? '?'
      // Determine reason: check future availability
      const futureCandidates = availableSlotKeys.filter((fs) => {
        if (um.absentDate) {
          const [fd] = fs.split('_')
          if (fd <= um.absentDate) return false
        }
        if (getSlotNumber(fs) === 0) return false
        return true
      })
      const studentAvail = futureCandidates.filter((fs) => isStudentAvailable(student, fs))
      const teacherAvail = futureCandidates.filter((fs) => hasAvailability(data.availability, instructorPersonType, um.teacherId, fs))
      const bothAvail = futureCandidates.filter((fs) =>
        isStudentAvailable(student, fs) && hasAvailability(data.availability, instructorPersonType, um.teacherId, fs)
      )
      let reason: string
      if (bothAvail.length === 0) {
        if (studentAvail.length === 0 && teacherAvail.length === 0) {
          reason = `講師(${teacherName})と生徒の出席可能日なし`
        } else if (studentAvail.length === 0) {
          reason = '生徒の出席可能日なし'
        } else if (teacherAvail.length === 0) {
          reason = `講師(${teacherName})の出席可能日なし`
        } else {
          reason = `講師(${teacherName})と生徒のマッチする日なし`
        }
      } else {
        // Both available but still not placed — check specific constraints
        const specificReasons: string[] = []
        const deskLimit = data.settings.deskCount ?? 0
        for (const fs of bothAvail) {
          const slotAssigns = nextAssignments[fs] ?? []
          const reasons: string[] = []
          // Desk count limit
          if (deskLimit > 0 && slotAssigns.length >= deskLimit) {
            reasons.push('机数上限')
          }
          // Teacher already used in slot
          if (slotAssigns.some((a) => a.teacherId === um.teacherId)) {
            const teacherAssign = slotAssigns.find((a) => a.teacherId === um.teacherId)
            if (teacherAssign && teacherAssign.studentIds.length >= 2) {
              reasons.push(`講師(${teacherName})のペアが満席`)
            } else if (teacherAssign) {
              // Teacher has space but student not placed — check student constraints
              if (teacherAssign.studentIds.some((sid) => constraintFor(data.constraints, sid, um.studentId) === 'incompatible')) {
                reasons.push('既存生徒との相性不可')
              }
            }
          } else {
            // Teacher not in slot — would need new pair
            if (deskLimit > 0 && slotAssigns.length >= deskLimit) {
              reasons.push('机数上限で新規ペア不可')
            }
          }
          // Teacher-student incompatible
          if (constraintFor(data.constraints, um.teacherId, um.studentId) === 'incompatible') {
            reasons.push(`講師(${teacherName})と生徒の相性不可`)
          }
          // Student already in another pair in this slot
          if (slotAssigns.some((a) => a.studentIds.includes(um.studentId))) {
            reasons.push('同コマに既に割当済み')
          }
          // 1日コマ上限 (per-student constraint card)
          const [fsDate] = fs.split('_')
          const slotsOnDate = Object.keys(nextAssignments).filter((k) => k.startsWith(fsDate + '_')).reduce((count, k) => 
            count + (nextAssignments[k]?.some((a) => a.studentIds.includes(um.studentId)) ? 1 : 0), 0)
          const umStudent = data.students.find((s) => s.id === um.studentId)
          const umCards = umStudent?.constraintCards ?? getDefaultConstraintCards(umStudent?.grade ?? '')
          const dailyLimit = umCards.includes('oneSlotOnly') ? 1 : umCards.includes('threeSlotLimit') ? 3 : 2
          if (slotsOnDate >= dailyLimit) {
            reasons.push(`1日${dailyLimit}コマ上限`)
          }
          if (reasons.length > 0) {
            for (const r of reasons) {
              if (!specificReasons.includes(r)) specificReasons.push(r)
            }
          }
        }
        reason = specificReasons.length > 0
          ? `講師(${teacherName}): ${specificReasons.join(', ')}`
          : `講師(${teacherName})の空き枠あるが配置されず(原因不明)`
      }
      const pendingDemand = pendingMakeupDemands.find((demand) =>
        demand.studentId === um.studentId
        && demand.teacherId === um.teacherId
        && demand.subject === um.subject
        && demand.absentDate === um.absentDate,
      )
      const analysis = teacher && pendingDemand
        ? analyzePlacementOptions(
            nextAssignments,
            effectiveForCounting,
            availableSlotKeys,
            student,
            um.subject,
            [teacher],
            {
              slotFilter: (fs) => !um.absentDate || fs.split('_')[0] > um.absentDate,
              makeupDemand: pendingDemand,
            },
          )
        : { force: [], teacher: [], cards: [], student: [], blockers: [] }
      const proposals = dedupeStatusProposals([
        ...analysis.force,
        ...analysis.teacher,
        ...analysis.student,
        ...(analysis.force.length === 0 ? analysis.cards : []),
      ])
      const causes = [reason, ...analysis.blockers]
      if (proposals.length > 0) {
        reason = `${reason} / ${proposals.map((proposal) => proposal.label).join(' / ')}`
      } else if (analysis.blockers.length > 0) {
        reason = `${reason} / 提案候補なし: ${analysis.blockers.join(', ')}`
      } else {
        reason = `${reason} / 提案候補なし: 机数上限・相性不可・出席可能日を確認してください`
      }
      unplacedMakeupEntries.push({ studentName: student.name, subject: um.subject, causes, proposals, reason })
    }

    // Final cleanup: remove teacher-only assignments (no valid students) that may have leaked through
    const validStudentIdSet2 = new Set(data.students.map((s) => s.id))
    for (const slot of Object.keys(nextAssignments)) {
      nextAssignments[slot] = nextAssignments[slot].filter((a) => a.studentIds.some(sid => sid && validStudentIdSet2.has(sid)) || a.isGroupLesson)
      if (nextAssignments[slot].length === 0) delete nextAssignments[slot]
    }

    await updateAssignments((current) => {
      // Merge: auto-assign results for non-recorded slots + preserve current assignments for recorded slots
      const mergedAssignments = { ...nextAssignments }
      for (const slot of recordedSlots) {
        mergedAssignments[slot] = current.assignments[slot] ?? []
      }
      return {
        ...current,
        assignments: mergedAssignments,
        autoAssignHighlights: { added: highlightAdded, changed: highlightChanged, makeup: highlightMakeup, changeDetails: highlightDetails, unplacedMakeup: unplacedMakeup.map((um) => ({ studentId: um.studentId, teacherId: um.teacherId, subject: um.subject, reason: unplacedMakeupEntries.find((e) => e.studentName === (data.students.find((s) => s.id === um.studentId)?.name ?? ''))?.reason ?? '' })) },
        settings: { ...current.settings, lastAutoAssignedAt: Date.now() },
      }
    })
    const overRemovedEntries = changeLog.filter((item) => item.action === '過割当解除' || item.action === '希望減で解除')
    const underSection: StatusSection = {
      key: 'under',
      title: '残コマあり',
      items: remainingStudents.flatMap((item) =>
        item.remaining.map((r) => ({
          label: `${item.studentName}: ${r.subject}残${r.remaining}コマ`,
          causes: r.suggestions.blockers.length > 0 ? r.suggestions.blockers : ['希望コマ数に対して割当が不足'],
          proposals: dedupeStatusProposals([
            ...r.suggestions.force,
            ...r.suggestions.teacher,
            ...r.suggestions.student,
            ...(r.suggestions.force.length === 0 ? r.suggestions.cards : []),
          ]),
        })),
      ),
    }
    const makeupSection: StatusSection = {
      key: 'makeup',
      title: '振替未配置',
      items: unplacedMakeupEntries.map((item) => ({
        label: `${item.studentName}: ${item.subject}`,
        causes: item.causes.length > 0 ? item.causes : [item.reason],
        proposals: item.proposals,
      })),
    }
    const overSection: StatusSection = {
      key: 'over',
      title: '過割当',
      items: overAssignedStudents.flatMap((item) =>
        item.over.map((r) => {
          const student = data.students.find((s) => s.name === item.studentName)
          const matchedSlots = student
            ? Object.entries(effectiveForCounting)
                .filter(([, slotAssignments]) => slotAssignments.some((a) => a.studentIds.includes(student.id) && getStudentSubject(a, student.id) === r.subject && !a.isRegular && !a.regularMakeupInfo?.[student.id]))
                .map(([slot]) => slotLabel(slot, isMendan, mendanStart))
            : []
          return {
            label: `${item.studentName}: ${r.subject} +${r.over}コマ`,
            causes: [`希望コマ数より ${r.over}コマ多く割り当て済み`],
            proposals: matchedSlots.length > 0
              ? [toStatusProposal(`後ろのコマから調整候補: ${matchedSlots.slice(-r.over).join(', ')}`)]
              : [toStatusProposal('自動提案を再実行するか、該当科目の割当を手動で減らしてください')],
          }
        }),
      ),
    }
    const shortageSection: StatusSection = {
      key: 'shortage',
      title: `${instructorLabel}不足`,
      items: shortageEntries.map((item) => ({
        label: slotLabel(item.slot, isMendan, mendanStart),
        causes: [item.detail],
        proposals: item.detail.includes('出席不可')
          ? [toStatusProposal(`${slotLabel(item.slot, isMendan, mendanStart)} に対応講師の出席可能コマを追加してください`)]
          : item.detail.includes('担当外科目')
            ? [toStatusProposal('担当可能な講師へ差し替えるか、対応講師の出席可能コマを追加してください')]
            : [toStatusProposal('講師設定と出席希望を確認してください')],
      })),
    }
    const overRemovedSection: StatusSection = {
      key: 'overRemoved',
      title: '過割当解除',
      items: overRemovedEntries.map((item) => ({
        label: slotLabel(item.slot, isMendan, mendanStart),
        causes: item.detail.split('\n').filter(Boolean),
        proposals: [toStatusProposal('必要なら希望コマ数や実績を見直してください')],
      })),
    }
    const report: StatusReport = {
      title: '自動提案結果',
      summary: changeLog.length > 0 ? `${changeLog.length}件の変更がありました。` : '変更はありませんでした。',
      sections: [underSection, makeupSection, overSection, shortageSection, overRemovedSection].filter((section) => section.items.length > 0),
    }
    setLatestStatusReport(report)
    setStatusModal(report)

    } catch (err) {
      console.error('[AutoAssign] Error:', err)
      alert(`自動割当エラー: ${err instanceof Error ? err.message : String(err)}`)
    } finally {
      setAutoAssignLoading(false)
    }
  }

  const resetAssignments = async (): Promise<void> => {
    if (!window.confirm('コマ割りをリセットしますか？\n（手動割当と自動提案結果を全てクリアします）')) return
    setManuallyModifiedSlots(new Set())
    await updateAssignments((current) => ({
      ...current,
      assignments: {},
      autoAssignHighlights: { added: {}, changed: {}, makeup: {} },
    }))
  }

  const getManualConstraintWarnings = (slot: string, assignment: Assignment, assignmentIdx: number): string[] => {
    if (!data || isMendan || !assignment.teacherId || assignment.studentIds.length === 0) {
      return []
    }

    const slotAssignments = data.assignments[slot] ?? []
    const warnings: string[] = []

    for (const studentId of assignment.studentIds) {
      const student = data.students.find((item) => item.id === studentId)
      if (!student) continue

      const reducedSlotAssignments = slotAssignments.flatMap((item, index) => {
        if (index !== assignmentIdx) return [item]

        const remainingIds = item.studentIds.filter((id) => id !== studentId)
        if (remainingIds.length === 0) return []

        const nextStudentSubjects = Object.entries(item.studentSubjects ?? {}).reduce<Record<string, string>>((acc, [id, subject]) => {
          if (id !== studentId) acc[id] = subject
          return acc
        }, {})
        const nextMakeupInfo = Object.entries(item.regularMakeupInfo ?? {}).reduce<Record<string, { dayOfWeek: number; slotNumber: number; date?: string }>>((acc, [id, info]) => {
          if (id !== studentId) acc[id] = info
          return acc
        }, {})

        return [{
          ...item,
          studentIds: remainingIds,
          ...(Object.keys(nextStudentSubjects).length > 0 ? { studentSubjects: nextStudentSubjects } : {}),
          ...(Object.keys(nextMakeupInfo).length > 0 ? { regularMakeupInfo: nextMakeupInfo } : {}),
        }]
      })

      const evalResult = evaluateConstraintCards(
        student,
        slot,
        { ...data.assignments, [slot]: reducedSlotAssignments },
        data.settings.slotsPerDay,
        data.regularLessons,
        data.groupLessons,
        assignment.teacherId,
      )

      if (evalResult.blocked) {
        warnings.push(`${student.name}: ${evalResult.blockReason ?? '制約カード違反'}`)
      }
    }

    return warnings
  }

  const setSlotTeacher = async (slot: string, idx: number, teacherId: string): Promise<void> => {
    await updateAssignments((current) => {
      const slotAssignments = [...(current.assignments[slot] ?? [])]
      if (!teacherId) {
        // Mark slot as manually modified so auto-fill won't recreate the deleted pair
        setManuallyModifiedSlots((prev) => new Set(prev).add(slot))
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
          // Manual assignment: allow any subject this teacher can handle for the student's grade.
          const student = current.students.find((s) => s.id === sid)
          const viable = student ? teachableBaseSubjects(teacher?.subjects ?? [], student.grade) : []
          newStudentSubjects[sid] = viable[0] ?? assignment.subject
        } else {
          // Manual assignment: keep any subject the teacher can still handle.
          const existingSubj = prevStudentSubjects[sid] ?? assignment.subject
          const student = current.students.find((s) => s.id === sid)
          const teacherCanTeach = student ? canTeachSubject(teacher?.subjects ?? [], student.grade, existingSubj) : false
          if (teacherCanTeach) {
            newStudentSubjects[sid] = existingSubj
          } else {
            const viable = student ? teachableBaseSubjects(teacher?.subjects ?? [], student.grade) : []
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
        studentSubjects: studentIds.length > 0 ? newStudentSubjects : {},
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
        slotAssignments[idx] = { ...assignment, subject, studentSubjects: assignment.studentIds.length > 0 ? studentSubjects : {} }
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
      if (!moved || moved.isGroupLesson) return current

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

      // When moving a regular assignment, convert to non-regular with regularMakeupInfo
      const movedCopy = { ...moved }
      if (movedCopy.isRegular) {
        const srcDayOfWeek = getSlotDayOfWeek(sourceSlot)
        const srcSlotNum = getSlotNumber(sourceSlot)
        const regularMakeupInfo: Record<string, { dayOfWeek: number; slotNumber: number }> = { ...(movedCopy.regularMakeupInfo ?? {}) }
        for (const sid of movedCopy.studentIds) {
          if (!regularMakeupInfo[sid]) {
            regularMakeupInfo[sid] = { dayOfWeek: srcDayOfWeek, slotNumber: srcSlotNum }
          }
        }
        movedCopy.isRegular = false
        movedCopy.regularMakeupInfo = regularMakeupInfo
      }

      // Move
      srcAssignments.splice(sourceIdx, 1)
      targetAssignments.push(movedCopy)
      const nextAssignments = { ...current.assignments }
      if (srcAssignments.length === 0) {
        delete nextAssignments[sourceSlot]
      } else {
        nextAssignments[sourceSlot] = srcAssignments
      }
      nextAssignments[targetSlot] = targetAssignments

      // Highlight moved pair as UPDATE
      const hl = current.autoAssignHighlights ?? {}
      const changedSigs = { ...(hl.changed ?? {}) }
      const sig = assignmentSignature(moved)
      changedSigs[targetSlot] = [...(changedSigs[targetSlot] ?? []), sig]
      // Remove old highlight from source slot
      if (changedSigs[sourceSlot]) {
        changedSigs[sourceSlot] = changedSigs[sourceSlot].filter((s) => s !== sig)
        if (changedSigs[sourceSlot].length === 0) delete changedSigs[sourceSlot]
      }
      const addedSigs = { ...(hl.added ?? {}) }
      if (addedSigs[sourceSlot]) {
        addedSigs[sourceSlot] = addedSigs[sourceSlot].filter((s) => s !== sig)
        if (addedSigs[sourceSlot].length === 0) delete addedSigs[sourceSlot]
      }
      const details = { ...(hl.changeDetails ?? {}) }
      if (!details[targetSlot]) details[targetSlot] = {}
      details[targetSlot][sig] = `ペア移動: ${sourceSlot} → ${targetSlot}`

      return { ...current, assignments: nextAssignments, autoAssignHighlights: { ...hl, changed: changedSigs, added: addedSigs, changeDetails: details } }
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
      if (!srcAssignment || srcAssignment.isGroupLesson) return current
      if (!srcAssignment.studentIds.includes(studentId)) return current

      // Check student/parent availability in target slot
      const student = current.students.find((s) => s.id === studentId)
      if (student && (isMendan
        ? !isParentAvailableForMendan(current.availability, student.id, targetSlot)
        : !isStudentAvailable(student, targetSlot))) return current

      // Check student not already in target slot (when same slot, exclude the source assignment being modified)
      const isSameSlotMove = sourceSlot === targetSlot
      const targetAssignments = [...(current.assignments[targetSlot] ?? [])]
      if (targetAssignments.some((a, aIdx) => {
        if (isSameSlotMove && aIdx === sourceIdx) return false // source will be modified
        return a.studentIds.includes(studentId)
      })) return current

      // Determine if this student is a regular student being moved (needs regularMakeupInfo)
      const isRegularSource = srcAssignment.isRegular
      const existingMakeupInfo = srcAssignment.regularMakeupInfo?.[studentId]
      const studentMakeupInfo = existingMakeupInfo ?? (isRegularSource ? { dayOfWeek: getSlotDayOfWeek(sourceSlot), slotNumber: getSlotNumber(sourceSlot), date: sourceSlot.split('_')[0] } : undefined)

      // Remove student from source assignment
      const updatedSrcStudentIds = srcAssignment.studentIds.filter((sid) => sid !== studentId)
      const updatedSrcStudentSubjects = { ...srcAssignment.studentSubjects }
      delete updatedSrcStudentSubjects[studentId]

      const studentSubject = getStudentSubject(srcAssignment, studentId)

      if (targetIdx !== undefined && targetIdx >= 0 && targetIdx < targetAssignments.length) {
        // Add to existing target assignment (max 2 students)
        const targetAssignment = targetAssignments[targetIdx]
        if (targetAssignment.studentIds.length >= 2) return current
        if (targetAssignment.isGroupLesson) return current

        const updatedTargetStudentIds = [...targetAssignment.studentIds, studentId]
        const updatedTargetStudentSubjects = { ...(targetAssignment.studentSubjects ?? {}) }
        // Set the student's subject in the target assignment
        updatedTargetStudentSubjects[studentId] = studentSubject
        // Carry over regularMakeupInfo for transferred regular students
        const updatedTargetMakeupInfo = studentMakeupInfo
          ? { ...(targetAssignment.regularMakeupInfo ?? {}), [studentId]: studentMakeupInfo }
          : targetAssignment.regularMakeupInfo
        targetAssignments[targetIdx] = {
          ...targetAssignment,
          studentIds: updatedTargetStudentIds,
          studentSubjects: updatedTargetStudentSubjects,
          isRegular: false, // Mark as manual so auto-fill won't overwrite
          ...(updatedTargetMakeupInfo ? { regularMakeupInfo: updatedTargetMakeupInfo } : {}),
        }
      } else {
        // Create new assignment in target slot with just this student
        const deskCount = current.settings.deskCount ?? 0
        if (deskCount > 0 && targetAssignments.length >= deskCount) return current

        // Auto-assign a compatible teacher who is available and not already used in this slot
        const usedTeacherIdsInTarget = new Set(targetAssignments.map((a) => a.teacherId).filter(Boolean))
        const instructorsList = isMendan ? current.managers : current.teachers
        const pType: PersonType = isMendan ? 'manager' : 'teacher'
        let autoTeacherId = ''
        // For regular students, prefer their regular teacher (only if they can teach the subject)
        const regLesson = !isMendan ? current.regularLessons.find(r => r.studentIds.includes(studentId)) : undefined
        if (regLesson?.teacherId && !usedTeacherIdsInTarget.has(regLesson.teacherId) && hasAvailability(current.availability, pType, regLesson.teacherId, targetSlot)) {
          const regTeacher = current.teachers.find(t => t.id === regLesson.teacherId)
          if (regTeacher && canTeachSubject(regTeacher.subjects, student?.grade ?? '', studentSubject)) {
            autoTeacherId = regLesson.teacherId
          }
        }
        if (!autoTeacherId) {
          for (const inst of instructorsList) {
            if (usedTeacherIdsInTarget.has(inst.id)) continue
            if (!hasAvailability(current.availability, pType, inst.id, targetSlot)) continue
            // Check subject compatibility (skip for mendan)
            if (!isMendan && student) {
              const compatible = canTeachSubject((inst as Teacher).subjects ?? [], student.grade, studentSubject)
              if (!compatible) continue
            }
            autoTeacherId = inst.id
            break
          }
        }

        targetAssignments.push({
          teacherId: autoTeacherId,
          studentIds: [studentId],
          subject: studentSubject,
          studentSubjects: { [studentId]: studentSubject },
          ...(studentMakeupInfo ? { regularMakeupInfo: { [studentId]: studentMakeupInfo } } : {}),
        })
      }

      // Update source: if no students left, remove the assignment entirely
      if (updatedSrcStudentIds.length === 0) {
        srcAssignments.splice(sourceIdx, 1)
      } else {
        // Remove moved student's regularMakeupInfo but keep isRegular for remaining students
        const updatedSrcMakeupInfo = { ...(srcAssignment.regularMakeupInfo ?? {}) }
        delete updatedSrcMakeupInfo[studentId]
        srcAssignments[sourceIdx] = {
          ...srcAssignment,
          studentIds: updatedSrcStudentIds,
          studentSubjects: updatedSrcStudentSubjects,
          // Keep isRegular if the source was regular — remaining students are still regular
          ...(Object.keys(updatedSrcMakeupInfo).length > 0 ? { regularMakeupInfo: updatedSrcMakeupInfo } : {}),
        }
      }

      // Mark both source and target slots as manually modified
      setManuallyModifiedSlots((prev) => {
        const next = new Set(prev)
        next.add(sourceSlot)
        next.add(targetSlot)
        return next
      })

      const nextAssignments = { ...current.assignments }
      if (isSameSlotMove) {
        // Same slot: merge source and target changes into one array
        const mergedAssignments = [...targetAssignments]
        // Apply source changes (student removed) to the merged array
        if (updatedSrcStudentIds.length === 0) {
          mergedAssignments.splice(sourceIdx, 1)
        } else {
          const updatedSrcMakeupInfo2 = { ...(srcAssignment.regularMakeupInfo ?? {}) }
          delete updatedSrcMakeupInfo2[studentId]
          mergedAssignments[sourceIdx] = {
            ...srcAssignment,
            studentIds: updatedSrcStudentIds,
            studentSubjects: updatedSrcStudentSubjects,
            ...(Object.keys(updatedSrcMakeupInfo2).length > 0 ? { regularMakeupInfo: updatedSrcMakeupInfo2 } : {}),
          }
        }
        nextAssignments[sourceSlot] = mergedAssignments
      } else {
        if (srcAssignments.length === 0) {
          delete nextAssignments[sourceSlot]
        } else {
          nextAssignments[sourceSlot] = srcAssignments
        }
        nextAssignments[targetSlot] = targetAssignments
      }

      // Highlight moved student's target assignment as UPDATE
      // For same-slot moves, use the merged assignments; for cross-slot, use targetAssignments
      const finalTargetAssignments = isSameSlotMove ? (nextAssignments[targetSlot] ?? []) : targetAssignments
      const movedTargetAssignment = finalTargetAssignments.find((a, aIdx) => {
        if (isSameSlotMove && aIdx === sourceIdx && updatedSrcStudentIds.length > 0) return false // skip modified source
        return a.studentIds.includes(studentId)
      })
      if (movedTargetAssignment) {
        const hl = current.autoAssignHighlights ?? {}
        const sig = assignmentSignature(movedTargetAssignment)
        const details = { ...(hl.changeDetails ?? {}) }
        if (!details[targetSlot]) details[targetSlot] = {}
        const studentName = student?.name ?? studentId
        // Regular student move → 振替 badge, otherwise → 変更 badge
        if (studentMakeupInfo) {
          const makeupSigs = { ...(hl.makeup ?? {}) }
          makeupSigs[targetSlot] = [...(makeupSigs[targetSlot] ?? []), sig]
          details[targetSlot][sig] = `振替: ${studentName} (${sourceSlot} → ${targetSlot})`
          return { ...current, assignments: nextAssignments, autoAssignHighlights: { ...hl, makeup: makeupSigs, changeDetails: details } }
        }
        const changedSigs = { ...(hl.changed ?? {}) }
        changedSigs[targetSlot] = [...(changedSigs[targetSlot] ?? []), sig]
        details[targetSlot][sig] = `生徒移動: ${studentName} (${sourceSlot} → ${targetSlot})`
        return { ...current, assignments: nextAssignments, autoAssignHighlights: { ...hl, changed: changedSigs, changeDetails: details } }
      }

      return { ...current, assignments: nextAssignments }
    })
  }

  // --- Save and close: create backup, then navigate to home ---
  const handleSaveAndClose = async () => {
    if (!classroomId) return
    const sessionName = data?.settings?.name ?? sessionId ?? ''
    try {
      await createBackup(classroomId, 'auto', [`セッション編集 (${sessionName})`])
      await cleanupOldBackups(classroomId, 30)
    } catch (e) {
      console.warn('[SaveAndClose] Backup failed:', e)
    }
    navigate(`/c/${classroomId}`, { state: { directHome: true } })
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
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
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
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
        </div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <div className="panel">
        <div className="admin-header-row">
          <h2 style={{ margin: 0 }}>管理画面: {data.settings.name} ({sessionId})</h2>
          <button className="btn btn-primary" type="button" onClick={() => void handleSaveAndClose()}>保存して閉じる</button>
        </div>
        <div className="admin-submission-row">
          <span className="muted" style={{ fontWeight: 600 }}>提出期間:</span>
          {data.settings.submissionStartDate && data.settings.submissionEndDate ? (
            <span className="badge ok">{data.settings.submissionStartDate} 〜 {data.settings.submissionEndDate}</span>
          ) : data.settings.submissionStartDate ? (
            <span className="badge warn">{data.settings.submissionStartDate} 〜 未設定</span>
          ) : (
            <span className="badge" style={{ background: '#f1f5f9', color: '#64748b' }}>未設定（制限なし）</span>
          )}
          <button className="btn secondary" type="button" style={{ padding: '2px 10px', fontSize: '12px' }}
            onClick={() => setShowSubmissionPicker((v) => !v)}>
            {showSubmissionPicker ? '閉じる' : '📅 変更'}
          </button>
          <span className="muted" style={{ fontSize: '11px' }}>※未設定＝制限なし。期間後もカレンダー表示あり</span>
        </div>
        {showSubmissionPicker && (
          <div style={{ marginTop: 8 }}>
            <DateRangePicker
              startDate={data.settings.submissionStartDate ?? ''}
              endDate={data.settings.submissionEndDate ?? ''}
              onStartChange={(d) => void update((c) => ({ ...c, settings: { ...c.settings, submissionStartDate: d || undefined } }))}
              onEndChange={(d) => void update((c) => ({ ...c, settings: { ...c.settings, submissionEndDate: d || undefined } }))}
            />
          </div>
        )}
        <p className="muted" style={{ marginTop: 6 }}>管理者のみ編集できます。希望入力は個別URLで配布してください。</p>
      </div>

      {!authorized ? (
        <div className="panel">
          <h3>管理者パスワードが一致しません</h3>
          <p className="muted">
            トップ画面で管理者パスワードを入力し「続行」してから、もう一度この特別講習を開いてください。
          </p>
          <Link to={`/c/${classroomId}`}>トップへ戻る</Link>
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
            <button className="btn secondary" type="button" style={{ marginTop: 8, fontSize: '0.85em' }}
              onClick={() => { if (confirm(`全${instructorLabel}（${instructors.length}名）のデータをランダム生成しますか？`)) void bulkRandomInstructors() }}>
              🎲 全{instructorLabel}一括ランダム入力 (DEV)
            </button>
          </div>

          <div className="panel">
            <h3>{isMendan ? '保護者一覧' : '生徒一覧'}</h3>
            <p className="muted">{isMendan ? '面談可能時間帯は保護者が希望URLから入力します。' : '希望コマ数・不可日は生徒本人が希望URLから入力します。'}</p>
            <table className="table">
              <thead><tr><th>名前</th>{!isMendan && <th>学年</th>}{!isMendan && <th>コマ制約</th>}<th>提出データ</th><th>代行入力</th><th>共有</th></tr></thead>
              <tbody>
                {data.students.map((student) => {
                  const cards = student.constraintCards ?? getDefaultConstraintCards(student.grade)
                  const isEditing = constraintEditStudentId === student.id
                  return (
                  <tr key={student.id}>
                    <td>
                      {student.name}{isMendan ? ' 保護者' : ''}
                      {recentlyUpdated.has(student.id) && (
                        <span className="badge ok" style={{ marginLeft: '8px', fontSize: '11px', animation: 'fadeIn 0.3s' }}>✓ 更新済</span>
                      )}
                    </td>
                    {!isMendan && <td>{student.grade}</td>}
                    {!isMendan && (
                      <td>
                        {cards.length > 0 && !isEditing && (
                          <span style={{ fontSize: '0.8em', color: '#475569', marginRight: 6 }}>
                            {summarizeConstraintCards(cards)}
                          </span>
                        )}
                        {isEditing ? (
                          <div style={{ minWidth: 280 }}>
                            {ALL_CONSTRAINT_CARDS.map((cardType) => {
                              const isChecked = cards.includes(cardType)
                              const conflictGroups = [CONSTRAINT_CARD_CONFLICT_GROUP, DAILY_LIMIT_CONFLICT_GROUP]
                              const isConflicting = !isChecked && conflictGroups.some((group) =>
                                group.includes(cardType) && cards.some((c) => group.includes(c) && c !== cardType)
                              )
                              return (
                                <label key={cardType} style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 4, fontSize: '0.85em', opacity: isConflicting ? 0.5 : 1, cursor: isConflicting ? 'not-allowed' : 'pointer' }}
                                  title={isConflicting ? '競合するカードが選択済みです' : CONSTRAINT_CARD_DESCRIPTIONS[cardType]}>
                                  <input type="checkbox" checked={isChecked} disabled={isConflicting}
                                    onChange={(e) => {
                                      let updated: ConstraintCardType[]
                                      if (e.target.checked) {
                                        // Remove conflicting cards from all conflict groups this card belongs to
                                        let withoutConflicts = [...cards]
                                        for (const group of conflictGroups) {
                                          if (group.includes(cardType)) {
                                            withoutConflicts = withoutConflicts.filter((c) => !group.includes(c))
                                          }
                                        }
                                        updated = [...withoutConflicts, cardType]
                                      } else {
                                        updated = cards.filter((c) => c !== cardType)
                                      }
                                      void update((cur) => ({
                                        ...cur,
                                        students: cur.students.map((s) => s.id === student.id ? { ...s, constraintCards: updated } : s),
                                      }))
                                    }} />
                                  <span>{CONSTRAINT_CARD_LABELS[cardType]}</span>
                                </label>
                              )
                            })}
                            {(() => {
                              const warnings = validateConstraintCards(cards)
                              return warnings.length > 0 ? (
                                <div style={{ color: '#dc2626', fontSize: '0.8em', margin: '4px 0', padding: '4px 8px', background: '#fef2f2', borderRadius: 4 }}>
                                  {warnings.map((w, i) => <div key={i}>⚠ {w}</div>)}
                                </div>
                              ) : null
                            })()}
                            <button type="button" className="btn secondary" style={{ fontSize: '0.75em', marginTop: 4 }}
                              onClick={() => setConstraintEditStudentId(null)}>閉じる</button>
                          </div>
                        ) : (
                          <button type="button" className="btn secondary" style={{ fontSize: '0.75em', padding: '2px 8px' }}
                            onClick={() => setConstraintEditStudentId(student.id)}>
                            {cards.length > 0 ? '編集' : '+ 設定'}
                          </button>
                        )}
                      </td>
                    )}
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
                  )
                })}
              </tbody>
            </table>
            <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap', alignItems: 'center' }}>
              <button className="btn secondary" type="button" style={{ fontSize: '0.85em' }}
                onClick={() => { if (confirm(`全${isMendan ? '保護者' : '生徒'}（${data.students.length}名）のデータをランダム生成しますか？`)) void bulkRandomStudents() }}>
                🎲 全{isMendan ? '保護者' : '生徒'}一括ランダム入力 (DEV)
              </button>
              {!isMendan && (
                <button className="btn secondary" type="button" style={{ fontSize: '0.85em' }}
                  onClick={async () => {
                    try {
                      const items = await listSessionItems(classroomId)
                      const otherSessions = items.filter((s) => s.id !== sessionId)
                      if (otherSessions.length === 0) { alert('他の講習セッションがありません'); return }
                      const picked = prompt(`制約を引き継ぐ講習を番号で選択:\n${otherSessions.map((s, i) => `${i + 1}. ${s.name}`).join('\n')}`, '1')
                      if (!picked) return
                      const idx = parseInt(picked, 10) - 1
                      if (idx < 0 || idx >= otherSessions.length) { alert('無効な番号です'); return }
                      const sourceSession = await loadSession(classroomId, otherSessions[idx].id)
                      if (!sourceSession) { alert('セッションデータの読み込みに失敗しました'); return }
                      // Build a map of student id → constraintCards from the source session
                      const constraintMap = new Map<string, ConstraintCardType[]>()
                      for (const s of sourceSession.students) {
                        const cards = s.constraintCards ?? (s as Record<string, unknown>).slotConstraints as ConstraintCardType[] | undefined
                        if (cards && cards.length > 0) {
                          constraintMap.set(s.id, cards)
                        }
                      }
                      if (constraintMap.size === 0) { alert(`「${otherSessions[idx].name}」にはコマ制約が設定されていません`); return }
                      const matchCount = data.students.filter((s) => constraintMap.has(s.id)).length
                      if (matchCount === 0) { alert('一致する生徒がいませんでした'); return }
                      if (!confirm(`「${otherSessions[idx].name}」から${matchCount}名分のコマ制約を引き継ぎますか？\n（既存の制約は上書きされます）`)) return
                      void update((cur) => ({
                        ...cur,
                        students: cur.students.map((s) => {
                          const inherited = constraintMap.get(s.id)
                          if (!inherited) return s
                          return { ...s, constraintCards: [...inherited] }
                        }),
                      }))
                      alert(`${matchCount}名のコマ制約を引き継ぎました`)
                    } catch (e) {
                      alert(`エラー: ${e instanceof Error ? e.message : String(e)}`)
                    }
                  }}>
                  📋 前回の講習から制約を引き継ぐ
                </button>
              )}
            </div>
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

                // Use live editing results for the recording slot (so badge updates in real-time during editing)
                const liveActualResults = recordingSlot
                  ? { ...(data.actualResults ?? {}), [recordingSlot]: editingResults.map(({ _uid: _, ...rest }) => rest) }
                  : data.actualResults
                const effAssignments = buildEffectiveAssignments(data.assignments, liveActualResults)
                const liveActual = liveActualResults ?? {}
                const recordedSlotSet = new Set(Object.keys(liveActual))
                const studentsWithRemaining = data.students
                  .map((student) => {
                    const remaining = Object.entries(student.subjectSlots)
                      .map(([subj, desired]) => {
                        const assigned = countStudentSubjectLoad(effAssignments, student.id, subj)
                        return { subj, rem: desired - assigned }
                      })
                      .filter((r) => r.rem !== 0)
                    // Count slots where student was planned but missing from actual result (regular and makeup assignments)
                    // Track per-subject for display, and reasons for no available makeup slot
                    type NoMakeupReason = 'no_student' | 'no_teacher' | 'no_match'
                    type AbsenceEntry = { subject: string; reasons: NoMakeupReason[] }
                    const absences: AbsenceEntry[] = []
                    for (const sk of slotKeys) {
                      if (!(sk in liveActual)) continue
                      const planned = data.assignments[sk] ?? []
                      const wasPlanned = planned.some((a) => a.studentIds.includes(student.id))
                      // Also check regularLessons as fallback (data.assignments may have been corrupted by previous auto-assign)
                      const [skDate] = sk.split('_')
                      const skDow = getIsoDayOfWeek(skDate)
                      const skSlotNum = getSlotNumber(sk)
                      const wasRegularLesson = data.regularLessons.some((rl) =>
                        rl.studentIds.includes(student.id) && rl.dayOfWeek === skDow && rl.slotNumber === skSlotNum
                      )
                      if (!wasPlanned && !wasRegularLesson) continue
                      const isInActual = liveActual[sk].some((r: ActualResult) => r.studentIds.includes(student.id))
                      if (!isInActual) {
                        // Determine assignment info (from planned or regularLessons)
                        const origAssignment = planned.find((a) => a.studentIds.includes(student.id))
                        const isRegularAbsence = origAssignment?.isRegular || (!origAssignment && wasRegularLesson)
                        const isMakeupAbsence = origAssignment && !origAssignment.isRegular && !!origAssignment.regularMakeupInfo?.[student.id]
                        if (!isRegularAbsence && !isMakeupAbsence) continue
                        // Get subject from assignment or regularLesson
                        let absentSubj: string
                        let origTeacherId: string
                        if (origAssignment) {
                          absentSubj = getStudentSubject(origAssignment, student.id)
                          origTeacherId = origAssignment.teacherId
                        } else {
                          const rl = data.regularLessons.find((r) =>
                            r.studentIds.includes(student.id) && r.dayOfWeek === skDow && r.slotNumber === skSlotNum
                          )!
                          absentSubj = rl.studentSubjects?.[student.id] ?? rl.subject
                          origTeacherId = rl.teacherId
                        }
                        const [absentDate] = sk.split('_')
                        // Check future slots after the absence date for makeup availability
                        const futureCandidates = slotKeys.filter((fs) => {
                          if (recordedSlotSet.has(fs)) return false
                          const [fd] = fs.split('_')
                          if (fd <= absentDate) return false
                          if (getSlotNumber(fs) === 0) return false
                          return true
                        })
                        const studentAvailSlots = futureCandidates.filter((fs) => isStudentAvailable(student, fs))
                        const teacherAvailSlots = futureCandidates.filter((fs) => hasAvailability(data.availability, instructorPersonType, origTeacherId, fs))
                        const bothAvailSlots = futureCandidates.filter((fs) =>
                          isStudentAvailable(student, fs) && hasAvailability(data.availability, instructorPersonType, origTeacherId, fs)
                        )
                        const reasons: NoMakeupReason[] = []
                        if (bothAvailSlots.length === 0) {
                          if (studentAvailSlots.length === 0 && teacherAvailSlots.length === 0) {
                            reasons.push('no_match')
                          } else if (studentAvailSlots.length === 0) {
                            reasons.push('no_student')
                          } else if (teacherAvailSlots.length === 0) {
                            reasons.push('no_teacher')
                          } else {
                            reasons.push('no_match')
                          }
                        }
                        absences.push({ subject: absentSubj, reasons })
                      }
                    }
                    // Count successful makeup placements in non-recorded slots
                    const makeupPlacedBySubj: Record<string, number> = {}
                    for (const sk of slotKeys) {
                      if (recordedSlotSet.has(sk)) continue
                      const slotAssigns = effAssignments[sk] ?? []
                      for (const a of slotAssigns) {
                        if (a.regularMakeupInfo?.[student.id]) {
                          const subj = getStudentSubject(a, student.id)
                          makeupPlacedBySubj[subj] = (makeupPlacedBySubj[subj] ?? 0) + 1
                        }
                      }
                    }
                    // Deduct placed makeups from absences (covered absences are removed)
                    const coveredRemaining: Record<string, number> = { ...makeupPlacedBySubj }
                    const uncoveredAbsences: AbsenceEntry[] = []
                    for (const abs of absences) {
                      if ((coveredRemaining[abs.subject] ?? 0) > 0) {
                        coveredRemaining[abs.subject]--
                        continue // This absence is covered by a makeup placement
                      }
                      uncoveredAbsences.push(abs)
                    }
                    const missingBySubj: Record<string, number> = {}
                    const noMakeupReasons: NoMakeupReason[] = []
                    for (const abs of uncoveredAbsences) {
                      missingBySubj[abs.subject] = (missingBySubj[abs.subject] ?? 0) + 1
                      noMakeupReasons.push(...abs.reasons)
                    }
                    const missingTotal = Object.values(missingBySubj).reduce((a, b) => a + b, 0)
                    if (remaining.length === 0 && missingTotal === 0) return null
                    return { name: student.name, remaining, missingBySubj, noMakeupReasons }
                  })
                  .filter(Boolean) as { name: string; remaining: { subj: string; rem: number }[]; missingBySubj: Record<string, number>; noMakeupReasons: ('no_student' | 'no_teacher' | 'no_match')[] }[]

                const currentPendingMakeupDemands = collectPendingMakeupDemands(data.assignments, liveActual)
                const currentAvailableSlotKeys = slotKeys.filter((s) => !recordedSlotSet.has(s))

                // Merge unplacedMakeup from autoAssignHighlights into studentsWithRemaining
                const unresolvedPendingKeys = new Set(currentPendingMakeupDemands.map((demand) => `${demand.studentId}|${demand.teacherId}|${demand.subject}|${demand.absentDate ?? ''}`))
                const unplacedMakeupHighlights = (data.autoAssignHighlights?.unplacedMakeup ?? []).filter((item) =>
                  unresolvedPendingKeys.has(`${item.studentId}|${item.teacherId}|${item.subject}|`)
                  || [...unresolvedPendingKeys].some((key) => key.startsWith(`${item.studentId}|${item.teacherId}|${item.subject}|`)),
                )
                for (const um of unplacedMakeupHighlights) {
                  const studentName = data.students.find((s) => s.id === um.studentId)?.name
                  if (!studentName) continue
                  const existing = studentsWithRemaining.find((s) => s.name === studentName)
                  if (existing) {
                    existing.missingBySubj[um.subject] = (existing.missingBySubj[um.subject] ?? 0) + 1
                  } else {
                    studentsWithRemaining.push({ name: studentName, remaining: [], missingBySubj: { [um.subject]: 1 }, noMakeupReasons: [] })
                  }
                }

                const missingTotal = (s: { missingBySubj: Record<string, number> }) => Object.values(s.missingBySubj).reduce((a, b) => a + b, 0)
                const underAssigned = studentsWithRemaining.filter((s) => s.remaining.some((r) => r.rem > 0) || missingTotal(s) > 0)
                const overAssigned = studentsWithRemaining.filter((s) => s.remaining.some((r) => r.rem < 0))
                const teacherShortages = collectTeacherShortages(data, effAssignments)

                const underTooltip = underAssigned
                  .map((s) => {
                    const parts: string[] = []
                    const missingEntries = Object.entries(s.missingBySubj).filter(([, c]) => c > 0)
                    if (missingEntries.length > 0) {
                      parts.push(missingEntries.map(([subj, count]) => `通常${subj}残${count}コマ`).join(', '))
                    }
                    if (s.noMakeupReasons.length > 0) {
                      const reasonCounts = { no_student: 0, no_teacher: 0, no_match: 0 }
                      for (const r of s.noMakeupReasons) reasonCounts[r]++
                      const reasonParts: string[] = []
                      if (reasonCounts.no_student > 0) reasonParts.push(`生徒の出席可能日なし${reasonCounts.no_student > 1 ? `×${reasonCounts.no_student}` : ''}`)
                      if (reasonCounts.no_teacher > 0) reasonParts.push(`講師の出席可能日なし${reasonCounts.no_teacher > 1 ? `×${reasonCounts.no_teacher}` : ''}`)
                      if (reasonCounts.no_match > 0) reasonParts.push(`講師と生徒のマッチする日なし${reasonCounts.no_match > 1 ? `×${reasonCounts.no_match}` : ''}`)
                      parts.push(`振替不可: ${reasonParts.join(', ')}`)
                    }
                    const specials = s.remaining.filter((r) => r.rem > 0)
                    if (specials.length > 0) parts.push(specials.map((r) => `特別${r.subj}残${r.rem}コマ`).join(', '))
                    return `${s.name}: ${parts.join(', ')}`
                  })
                  .join('\n')
                const overTooltip = overAssigned
                  .map((s) => `${s.name}: ${s.remaining.filter((r) => r.rem < 0).map((r) => `${r.subj}${r.rem}`).join(', ')}`)
                  .join('\n')
                const shortageTooltip = teacherShortages
                  .map((item) => `${slotLabel(item.slot, isMendan, mendanStart)}: ${item.detail}`)
                  .join('\n')
                const currentSections: StatusSection[] = [
                  {
                    key: 'under' as const,
                    title: '残コマあり',
                    items: underAssigned.map((s) => {
                      const causes: string[] = []
                      const proposalPool: StatusProposal[] = []
                      const missingEntries = Object.entries(s.missingBySubj).filter(([, c]) => c > 0)
                      const currentStudent = data.students.find((student) => student.name === s.name)
                      if (missingEntries.length > 0) {
                        causes.push(...missingEntries.map(([subj, count]) => `通常${subj}残${count}コマ`))
                      }
                      const specials = s.remaining.filter((r) => r.rem > 0)
                      if (specials.length > 0) {
                        causes.push(...specials.map((r) => `特別${r.subj}残${r.rem}コマ`))
                      }
                      if (currentStudent) {
                        for (const special of specials) {
                          const analysis = buildRemainingSuggestions(data.assignments, effAssignments, currentAvailableSlotKeys, currentStudent, special.subj)
                          proposalPool.push(...analysis.force, ...analysis.teacher, ...analysis.student, ...(analysis.force.length === 0 ? analysis.cards : []))
                        }
                        const studentMakeups = currentPendingMakeupDemands.filter((demand) => demand.studentId === currentStudent.id)
                        for (const demand of studentMakeups) {
                          const teacher = data.teachers.find((item) => item.id === demand.teacherId)
                          if (!teacher) continue
                          const analysis = analyzePlacementOptions(
                            data.assignments,
                            effAssignments,
                            currentAvailableSlotKeys,
                            currentStudent,
                            demand.subject,
                            [teacher],
                            {
                              slotFilter: (slot) => !demand.absentDate || slot.split('_')[0] > demand.absentDate,
                              makeupDemand: demand,
                            },
                          )
                          proposalPool.push(...analysis.force, ...analysis.teacher, ...analysis.student, ...(analysis.force.length === 0 ? analysis.cards : []))
                        }
                      }
                      if (proposalPool.length === 0 && s.noMakeupReasons.includes('no_teacher')) proposalPool.push(toStatusProposal('講師の出席可能コマを増やしてください'))
                      if (proposalPool.length === 0 && s.noMakeupReasons.includes('no_student')) proposalPool.push(toStatusProposal('生徒の出席不可日・不可コマを減らしてください'))
                      if (proposalPool.length === 0 && s.noMakeupReasons.includes('no_match')) proposalPool.push(toStatusProposal('講師出席可能コマと生徒出席可能日が重なるように調整してください'))
                      const proposals = dedupeStatusProposals(proposalPool)
                      return {
                        label: s.name,
                        causes,
                        proposals: proposals.length > 0 ? proposals : [toStatusProposal('自動提案結果モーダルで個別候補を確認するか、該当生徒の制約カードと出席希望を見直してください')],
                      }
                    }),
                  },
                  {
                    key: 'over' as const,
                    title: '過割当',
                    items: overAssigned.map((s) => ({
                      label: s.name,
                      causes: s.remaining.filter((r) => r.rem < 0).map((r) => `${r.subj}${r.rem}`),
                      proposals: [toStatusProposal('後ろのコマから手動調整するか、自動提案を再実行して過割当解除を確認してください')],
                    })),
                  },
                  {
                    key: 'shortage' as const,
                    title: `${instructorLabel}不足`,
                    items: teacherShortages.map((item) => ({
                      label: slotLabel(item.slot, isMendan, mendanStart),
                      causes: [item.detail],
                      proposals: item.detail.includes('出席不可')
                        ? [toStatusProposal(`${slotLabel(item.slot, isMendan, mendanStart)} に対応${instructorLabel}の出席可能コマを追加してください`)]
                        : [toStatusProposal('担当可能な講師への差し替え、または出席可能コマの追加を検討してください')],
                    })),
                  },
                ].filter((section) => section.items.length > 0)
                const currentStatusReport: StatusReport = {
                  title: '現在のコマ割り状況',
                  summary: '全員割当以外の項目に対する原因と提案を表示しています。',
                  sections: currentSections,
                }

                const hasAnyAssignment = Object.keys(data.assignments).length > 0
                const hasAnyDesired = data.students.some((s) => Object.values(s.subjectSlots).some((v) => v > 0))

                return (
                  <>
                    {!hasAnyAssignment || !hasAnyDesired ? (
                      <span className="badge" style={{ background: '#e5e7eb', color: '#374151' }}>未割当</span>
                    ) : underAssigned.length > 0 ? (
                      <span className="badge warn" title={underTooltip} style={{ cursor: 'pointer' }} onClick={() => setStatusModal({ ...currentStatusReport, sections: currentStatusReport.sections.filter((s) => s.key === 'under') })}>
                        残コマあり: {underAssigned.length}名
                      </span>
                    ) : (
                      <span className="badge ok">全員割当完了</span>
                    )}
                    {overAssigned.length > 0 && (
                      <span className="badge" title={overTooltip} style={{ cursor: 'pointer', background: '#fef2f2', color: '#dc2626', border: '1px solid #fca5a5' }} onClick={() => setStatusModal({ ...currentStatusReport, sections: currentStatusReport.sections.filter((s) => s.key === 'over') })}>
                        過割当: {overAssigned.length}名
                      </span>
                    )}
                    {teacherShortages.length > 0 && (
                      <span className="badge" title={shortageTooltip} style={{ cursor: 'pointer', background: '#fff1f2', color: '#be123c', border: '1px solid #fda4af' }} onClick={() => setStatusModal({ ...currentStatusReport, sections: currentStatusReport.sections.filter((s) => s.key === 'shortage') })}>
                        {instructorLabel}不足: {teacherShortages.length}件
                      </span>
                    )}
                  </>
                )
              })()}
              <button className="btn secondary" type="button" onClick={() => void applyAutoAssign()} disabled={autoAssignLoading}>
                {isMendan ? '自動割当（先着順）' : '自動提案'}
              </button>
              {latestStatusReport && !autoAssignLoading && (
                <button className="btn secondary" type="button" onClick={() => setStatusModal(latestStatusReport)}>
                  結果詳細
                </button>
              )}
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
              {data.settings.confirmed && (
                <button
                  className={`btn${showSalary ? '' : ' secondary'}`}
                  type="button"
                  style={showSalary ? { background: '#7c3aed', borderColor: '#7c3aed' } : {}}
                  onClick={() => setShowSalary((v) => !v)}
                >
                  💰 給与計算
                </button>
              )}
            </div>
            <p className="muted">{isMendan ? 'マネージャー1人 + 保護者1人の面談を先着順で自動割当。' : '通常授業は日付確定時に自動配置。特別講習は自動提案で割当。講師1人 + 生徒1〜2人。'}</p>
            <p className="muted" style={{ fontSize: '12px' }}>{isMendan ? 'クリックで別コマへ移動可' : '★=通常授業生徒　⚠=制約不可　クリックでペア/生徒を別コマへ移動可'}</p>
            {/* Salary calculation panel */}
            {showSalary && (() => {
              const rates = data.tierRates ?? { ...defaultTierRates }
              const salaryRows = computeSalaryData(rates)
              const grandTotal = salaryRows.reduce((sum, r) => sum + r.total, 0)
              const totalAllSlots = salaryRows.reduce((sum, r) => sum + r.A + r.B + r.C + r.D, 0)
              const recordedCount = Object.keys(data.actualResults ?? {}).length
              const tierLabels: { key: 'A' | 'B' | 'C' | 'D'; label: string; desc: string }[] = [
                { key: 'A', label: 'A', desc: '生徒1人・中学生以下' },
                { key: 'B', label: 'B', desc: '生徒2人・両方中学生以下' },
                { key: 'C', label: 'C', desc: '生徒1人・高校生以上' },
                { key: 'D', label: 'D', desc: '生徒2人・片方以上が高校生' },
              ]
              return (
                <div style={{ background: '#f5f3ff', border: '1px solid #c4b5fd', borderRadius: '8px', padding: '16px', marginBottom: '12px' }}>
                  <h3 style={{ margin: '0 0 8px', fontSize: '16px' }}>💰 給与計算</h3>
                  <p style={{ fontSize: '0.8em', color: '#6b7280', margin: '0 0 10px' }}>実績記録済みコマ: {recordedCount} / {slotKeys.length}　※通常コマ（固定授業）は給与計算に含まれません</p>
                  {/* Global tier rate settings */}
                  <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', marginBottom: '12px', padding: '8px', background: '#ede9fe', borderRadius: '6px' }}>
                    {tierLabels.map((t) => (
                      <div key={t.key} style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <label style={{ fontSize: '0.8em', fontWeight: 'bold' }}>{t.key}<span style={{ fontSize: '0.85em', color: '#6b7280', fontWeight: 'normal' }}>({t.desc})</span></label>
                        <input
                          type="number"
                          min={0}
                          step={100}
                          value={rates[t.key]}
                          style={{ width: '80px', textAlign: 'right', fontSize: '0.9em' }}
                          onChange={(e) => void saveTierRate(t.key, Number(e.target.value))}
                        />
                        <span style={{ fontSize: '0.8em', color: '#6b7280' }}>円</span>
                      </div>
                    ))}
                  </div>
                  {salaryRows.length === 0 ? (
                    <p style={{ color: '#6b7280', fontSize: '0.9em' }}>実績が記録されていません。各コマの「📝 実績記録」ボタンから記録してください。</p>
                  ) : (
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '0.85em' }}>
                        <thead>
                          <tr style={{ borderBottom: '2px solid #c4b5fd' }}>
                            <th style={{ textAlign: 'left', padding: '6px 8px' }}>{isMendan ? 'マネージャー' : '講師'}</th>
                            {tierLabels.map((t) => (
                              <th key={t.key} style={{ textAlign: 'center', padding: '6px 4px' }} title={t.desc}>{t.key}</th>
                            ))}
                            <th style={{ textAlign: 'center', padding: '6px 4px' }}>計</th>
                            <th style={{ textAlign: 'right', padding: '6px 8px' }}>合計 (円)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {salaryRows.map((row) => (
                            <tr key={row.teacherId} style={{ borderBottom: '1px solid #e5e7eb' }}>
                              <td style={{ padding: '4px 8px', whiteSpace: 'nowrap' }}>{row.name}</td>
                              {tierLabels.map((t) => (
                                <td key={t.key} style={{ textAlign: 'center', padding: '4px 4px' }}>{row[t.key] || '—'}</td>
                              ))}
                              <td style={{ textAlign: 'center', padding: '4px 4px', fontWeight: 'bold' }}>{row.A + row.B + row.C + row.D}</td>
                              <td style={{ textAlign: 'right', padding: '4px 8px', fontWeight: 'bold' }}>{row.total.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr style={{ borderTop: '2px solid #c4b5fd', fontWeight: 'bold' }}>
                            <td style={{ padding: '6px 8px' }}>合計</td>
                            {tierLabels.map((t) => (
                              <td key={t.key} style={{ textAlign: 'center', padding: '4px 4px' }}>{salaryRows.reduce((s, r) => s + r[t.key], 0) || '—'}</td>
                            ))}
                            <td style={{ textAlign: 'center', padding: '4px 4px' }}>{totalAllSlots}</td>
                            <td style={{ textAlign: 'right', padding: '6px 8px' }}>{grandTotal.toLocaleString()}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  )}
                </div>
              )
            })()}
            {showRules && (
              <div className="rules-panel" style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '8px', padding: '16px 20px', marginBottom: '12px', fontSize: '14px', lineHeight: '1.8' }}>
                <h3 style={{ margin: '0 0 12px', fontSize: '16px' }}>📖 コマ割りルール</h3>
                {isMendan ? (
                  <div style={{ display: 'grid', gap: '8px' }}>
                    <p style={{ margin: 0, color: '#475569' }}>1コマ = マネージャー1人 ＋ 保護者1人。提出順（先着）で優先割当。</p>
                    <p style={{ margin: 0, color: '#475569' }}>マネージャーの空き ∩ 保護者の希望が一致するコマに配置。机数上限あり。</p>
                  </div>
                ) : (
                  <div style={{ display: 'grid', gap: '10px' }}>
                    <section>
                      <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>基本</h4>
                      <p style={{ margin: 0, color: '#475569' }}>1コマ = 講師1人 ＋ 生徒2人まで。★=通常授業生徒（マスタから自動配置）。■=集団授業。机数上限あり。</p>
                    </section>
                    <section>
                      <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>共通ルール（自動提案スコアリング・優先順）</h4>
                      <ol style={{ margin: 0, paddingLeft: '20px', color: '#475569', fontSize: '13px' }}>
                        <li><b>2人ペアボーナス +1000</b> → 講師稼働率の最大化</li>
                        <li><b>既出勤日に追加 +500</b> / 新規出勤日 −200 → 講師の出勤日数を最小化</li>
                        <li><b>残コマ多数優先 ×20</b> → 生徒競合の場合残コマ数が多い生徒を優先</li>
                        <li><b>生徒は一日上限2コマまで</b>（ハードフィルタ）</li>
                      </ol>
                    </section>
                    <section>
                      <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>コマ制約カード（共通ルールより優先）</h4>
                      <ul style={{ margin: 0, paddingLeft: '20px', color: '#475569', fontSize: '13px' }}>
                        <li><b>(デフォルト) 受験生以外の後半コマ優先</b> → 高3・中3以外は3限以降に配置しやすくし、2人ペアの形成を促進</li>
                        <li><b>(デフォルト) 集団後連続</b> → 集団授業がある日の中3は、午前の後に早いコマから2コマ連続で配置</li>
                        <li><b>通常講師優先</b> → 通常授業の講師を優先</li>
                        <li><b>2コマ連続</b> → 生徒を2コマ連続で配置する（複数科目の残コマがある場合、科目は前後で分ける）</li>
                        <li><b>2コマ連続(一コマ空け)</b> → 生徒を2コマ連続で配置するが、間に1コマ入れる</li>
                        <li><b>一コマ限定</b> → 生徒を1日1コマに限定する</li>
                        <li><b>通常授業連結</b> → 通常授業の前後に特別講習のコマをつなげ2コマ連続とする</li>
                      </ul>
                      <p style={{ margin: '4px 0 0', color: '#94a3b8', fontSize: '12px' }}>※ 2コマ連続 / 2コマ連続(一コマ空け) / 一コマ限定 / 通常授業連結 は競合（同一生徒に1つのみ選択可）</p>
                    </section>
                    <section>
                      <h4 style={{ margin: '0 0 4px', fontSize: '14px', color: '#334155' }}>制約</h4>
                      <p style={{ margin: 0, color: '#475569' }}>講師×生徒 / 講師×学年の不可制約（⚠マーク）。科目の共通性が必要。</p>
                    </section>
                  </div>
                )}
              </div>
            )}
            {showAnalytics && <AnalyticsPanel data={data} slotKeys={isMendan ? effectiveSlotKeys : slotKeys} />}
            <div style={{ position: 'relative' }}>
            {autoAssignLoading && (
              <div style={{ position: 'absolute', inset: 0, background: 'rgba(255,255,255,0.75)', zIndex: 20, display: 'flex', alignItems: 'flex-start', justifyContent: 'center', paddingTop: '18px', borderRadius: '8px' }}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '10px', background: '#fff', padding: '16px 28px', borderRadius: '10px', boxShadow: '0 2px 12px rgba(0,0,0,0.15)' }}>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', fontSize: '1em', color: '#334155', fontWeight: 600 }}>
                    <span className="spinner" style={{ width: 22, height: 22, borderWidth: 3 }} />
                    自動コマ割り実行中...
                  </span>
                  <div style={{ width: '200px', height: '6px', background: '#e2e8f0', borderRadius: '3px', overflow: 'hidden' }}>
                    <div style={{ width: `${Math.round(autoAssignProgress * 100)}%`, height: '100%', background: '#3b82f6', borderRadius: '3px', transition: 'width 0.2s ease' }} />
                  </div>
                </div>
              </div>
            )}
            <div className="grid-slots">
              {(isMendan ? effectiveSlotKeys : slotKeys).map((slot) => {
                const slotAssignments = data.assignments[slot] ?? []
                const isRecorded = data.actualResults != null && slot in data.actualResults
                // When actual results are recorded, show those instead of original assignments
                const displayAssignments = isRecorded ? (data.actualResults![slot] as Assignment[]) : slotAssignments
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
                // For student drag to a different slot: check if student can be placed (existing block accepts OR a compatible teacher is available for new pair)
                const studentDragNoTarget = (() => {
                  if (!isStudentDrag || isSameSlot) return false
                  const draggedStudent = data.students.find((s) => s.id === dragInfo.studentDragId)
                  const dragSubject = dragInfo.studentDragSubject ?? ''
                  const regTid = dragInfo.regularTeacherId
                  // Check if any existing block can accept (has space, not group, and teacher can teach subject)
                  const hasAcceptableBlock = slotAssignments.some((a) => {
                    if (a.isGroupLesson || a.studentIds.length >= 2) return false
                    // For regular student: only allow their regular teacher's block
                    if (regTid && a.teacherId !== regTid) return false
                    if (!isMendan && draggedStudent && dragSubject && a.teacherId) {
                      const teacher = instructors.find(t => t.id === a.teacherId) as Teacher | undefined
                      if (teacher && !canTeachSubject(teacher.subjects ?? [], draggedStudent.grade, dragSubject)) return false
                    }
                    return true
                  })
                  if (hasAcceptableBlock) return false
                  // No existing block accepts — check if a compatible teacher exists for new pair
                  if (regTid) {
                    // Regular student: only their regular teacher can form a new pair
                    if (usedTeacherIds.has(regTid)) return true
                    if (!hasAvailability(data.availability, instructorPersonType, regTid, slot)) return true
                    return false
                  }
                  for (const inst of instructors) {
                    if (usedTeacherIds.has(inst.id)) continue
                    if (!hasAvailability(data.availability, instructorPersonType, inst.id, slot)) continue
                    if (!isMendan && draggedStudent) {
                      if (!canTeachSubject(inst.subjects ?? [], draggedStudent.grade, dragSubject)) continue
                    }
                    return false // found at least one compatible teacher
                  }
                  return true // no compatible teacher found
                })()
                // For student drag to same slot: valid if dropping onto a different assignment within the same slot
                const isStudentDragSameSlotOk = isStudentDrag && isSameSlot
                const isDropValid = isDragActive && !isRecorded && (!isSameSlot || isStudentDragSameSlotOk) && !isDeskFull && !isTeacherConflict && !hasUnavailableStudent && !hasStudentConflict && !hasTeacherUnavailable && !studentDragNoTarget
                const slotDragClass = isDragActive ? (isSameSlot && !isStudentDragSameSlotOk ? '' : isDropValid ? ' drag-valid' : ' drag-invalid') : ''
                const isSourceSlot = isDragActive && dragInfo.sourceSlot === slot
                // For student drag: check if an unused compatible teacher exists for creating a new pair in this slot
                const hasUnusedCompatibleTeacher = (() => {
                  if (!isStudentDrag || isSameSlot) return false
                  const draggedStudent = data.students.find((s) => s.id === dragInfo.studentDragId)
                  const dragSubject = dragInfo.studentDragSubject ?? ''
                  const dCount = data.settings.deskCount ?? 0
                  if (dCount > 0 && slotAssignments.length >= dCount) return false
                  const regTid = dragInfo.regularTeacherId
                  if (regTid) {
                    // Regular student: only their regular teacher can form a new pair
                    if (usedTeacherIds.has(regTid)) return false
                    return hasAvailability(data.availability, instructorPersonType, regTid, slot)
                  }
                  for (const inst of instructors) {
                    if (usedTeacherIds.has(inst.id)) continue
                    if (!hasAvailability(data.availability, instructorPersonType, inst.id, slot)) continue
                    if (!isMendan && draggedStudent) {
                      if (!canTeachSubject(inst.subjects ?? [], draggedStudent.grade, dragSubject)) continue
                    }
                    return true
                  }
                  return false
                })()

                return (
                  <div className={`slot-card${slotDragClass}${isRecorded ? ' slot-recorded' : ''}`} key={slot}
                  >
                    <div className="slot-title">
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                        {slotLabel(slot, isMendan, mendanStart)}
                        {getSlotNumber(slot) !== 0 && (data.settings.deskCount ?? 0) > 0 && (
                          <span style={{ fontSize: '0.75em', color: slotAssignments.length >= (data.settings.deskCount ?? 0) ? '#dc2626' : '#6b7280' }}>
                            {slotAssignments.length}/{data.settings.deskCount}
                          </span>
                        )}
                        {isRecorded && <span style={{ fontSize: '0.7em', color: '#16a34a', fontWeight: 'bold' }}>✅ 実績済</span>}
                      </div>
                      {!isDragActive && (
                      <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                        {data.settings.confirmed && !isRecorded && (
                          <button
                            className="btn secondary"
                            type="button"
                            style={{ fontSize: '0.7em', padding: '2px 6px' }}
                            onClick={() => startRecording(slot)}
                          >
                            📝 実績記録
                          </button>
                        )}
                        {isRecorded && (
                          <button
                            className="btn secondary"
                            type="button"
                            style={{ fontSize: '0.7em', padding: '2px 6px' }}
                            onClick={() => startRecording(slot)}
                          >
                            📝 実績修正
                          </button>
                        )}
                        {isRecorded && recordingSlot !== slot && (
                          <button
                            className="btn secondary"
                            type="button"
                            style={{ fontSize: '0.7em', padding: '2px 6px', background: '#fee2e2', borderColor: '#fca5a5', color: '#dc2626' }}
                            onClick={() => void clearActualResults(slot)}
                          >
                            🔓 実績解除
                          </button>
                        )}
                        {(!isRecorded || recordingSlot === slot) && (
                          <button
                            className="btn secondary slot-add-btn"
                            type="button"
                            title="ペア追加"
                            onClick={() => recordingSlot === slot ? addEditingResultPair() : void addSlotAssignment(slot)}
                          >
                            ＋
                          </button>
                        )}
                      </div>
                      )}
                    </div>
                    {/* Click-move active: show cancel on source slot, destination button on valid target slots */}
                    {isDragActive && isSourceSlot && (
                      <div style={{ background: '#eff6ff', border: '1px solid #3b82f6', borderRadius: '6px', padding: '6px 8px', marginBottom: '4px', fontSize: '0.82em', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <span style={{ fontWeight: 600, color: '#1e40af' }}>
                          {isStudentDrag ? `${data.students.find(s => s.id === dragInfo.studentDragId)?.name ?? '?'} を移動中...` : 'ペアを移動中...'}
                        </span>
                        <button className="btn secondary" type="button" style={{ fontSize: '0.85em', padding: '2px 8px' }}
                          onClick={() => { setDragInfo(null); setTransferSlot(null) }}>キャンセル</button>
                      </div>
                    )}
                    {isDragActive && !isSourceSlot && isDropValid && (!isStudentDrag || hasUnusedCompatibleTeacher) && (
                      <button
                        className="btn"
                        type="button"
                        style={{ width: '100%', fontSize: '0.82em', padding: '4px', marginBottom: '4px', background: '#dcfce7', border: '1px solid #22c55e', color: '#15803d' }}
                        onClick={() => {
                          if (dragInfo.studentDragId) {
                            // Always create a new pair with auto-assigned compatible teacher
                            void moveStudentToSlot(dragInfo.sourceSlot, dragInfo.sourceIdx, dragInfo.studentDragId, slot)
                          } else {
                            void moveAssignment(dragInfo.sourceSlot, dragInfo.sourceIdx, slot)
                          }
                          setDragInfo(null)
                          setTransferSlot(null)
                        }}
                      >
                        このコマに移動
                      </button>
                    )}
                    {/* Actual result recording panel */}
                    {!isDragActive && recordingSlot === slot && (() => {
                      // Collect all student IDs used in this slot's editing results
                      const allUsedStudentIds = new Set(editingResults.flatMap((r) => r.studentIds.filter(Boolean)))
                      return (
                      <div className="recording-panel" style={{ background: '#fffbeb', border: '1px solid #f59e0b', borderRadius: '6px', padding: '10px', marginBottom: '4px' }}>
                        <div className="list">
                          {editingResults.map((result, rIdx) => (
                            <div key={result._uid ?? rIdx} className="assignment-block" style={{ position: 'relative' }}>
                              <button type="button" className="pair-delete-btn" title="このペアを削除"
                                onClick={() => removeEditingResultPair(rIdx)}>×</button>
                              <select value={result.teacherId}
                                onChange={(e) => updateEditingResult(rIdx, 'teacherId', e.target.value)}>
                                <option value="">{isMendan ? 'マネージャーを選択' : '講師を選択'}</option>
                                {instructors.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                              </select>
                              {result.teacherId && (
                                <>
                                  {(isMendan ? [0] : result.studentIds.map((_, i) => i)).map((pos) => {
                                    const sid = result.studentIds[pos] ?? ''
                                    const currentStudent = data.students.find((s) => s.id === sid)
                                    const selectedTeacherForResult = instructors.find((t) => t.id === result.teacherId)
                                    const studentSubjectOptions = selectedTeacherForResult && currentStudent
                                      ? teachableBaseSubjects(selectedTeacherForResult.subjects, currentStudent.grade).filter((subj) => currentStudent.subjects.includes(subj))
                                      : (selectedTeacherForResult ? [...new Set(selectedTeacherForResult.subjects.map(s => getSubjectBase(s)))] : [])
                                    // Hide students already used in other pairs in this slot
                                    const ownStudentIds = new Set(result.studentIds)
                                    const availableStudents = data.students.filter((s) =>
                                      ownStudentIds.has(s.id) || !allUsedStudentIds.has(s.id)
                                    )
                                    return (
                                      <div key={pos} className="student-select-row">
                                        <select value={sid}
                                          onChange={(e) => {
                                            const newIds = [...result.studentIds]
                                            newIds[pos] = e.target.value
                                            updateEditingResult(rIdx, 'studentIds', newIds)
                                          }}>
                                          <option value="">{isMendan ? '保護者を選択' : `生徒${pos + 1}を選択`}</option>
                                          {availableStudents.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                                        </select>
                                        {!isMendan && sid && studentSubjectOptions.length > 0 && (
                                          <select className="subject-select-inline"
                                            value={result.studentSubjects?.[sid] ?? result.subject ?? studentSubjectOptions[0] ?? ''}
                                            onChange={(e) => updateEditingResultStudentSubject(rIdx, sid, e.target.value)}>
                                            {studentSubjectOptions.map((subj) => (
                                              <option key={subj} value={subj}>{subj}</option>
                                            ))}
                                          </select>
                                        )}
                                        {sid && (
                                          <button type="button" className="student-clear-btn" title="この生徒を未選択にする"
                                            onClick={() => {
                                              const newIds = [...result.studentIds]
                                              newIds[pos] = ''
                                              updateEditingResult(rIdx, 'studentIds', newIds)
                                            }}>×</button>
                                        )}
                                      </div>
                                    )
                                  })}
                                  {!isMendan && result.studentIds.filter((s) => s).length < 2 && result.studentIds.length < 2 && (
                                    <button type="button" className="btn secondary" style={{ fontSize: '0.7em', padding: '2px 6px', marginTop: '4px' }}
                                      onClick={() => {
                                        const newIds = [...result.studentIds, '']
                                        updateEditingResult(rIdx, 'studentIds', newIds)
                                      }}>
                                      + 生徒追加
                                    </button>
                                  )}
                                </>
                              )}
                            </div>
                          ))}
                        </div>
                        <div style={{ display: 'flex', gap: '6px', marginTop: '8px', justifyContent: 'flex-end' }}>
                          <button type="button" className="btn" style={{ fontSize: '0.8em', background: '#16a34a', borderColor: '#16a34a' }} onClick={() => void saveActualResults()}>保存</button>
                          <button type="button" className="btn secondary" style={{ fontSize: '0.8em' }} onClick={cancelRecording}>キャンセル</button>
                        </div>
                      </div>
                      )
                    })()}
                    {recordingSlot !== slot && (
                    <div className="list">
                      {displayAssignments.map((assignment, idx) => {
                        // Student drag: hide assignment blocks that can't accept the student
                        if (isStudentDrag && !isSourceSlot) {
                          const canAccept = !assignment.isGroupLesson && assignment.studentIds.length < 2
                          if (!canAccept) return null
                        }
                        const selectedTeacher = instructors.find((t) => t.id === assignment.teacherId)

                        const isIncompatiblePair = !isMendan && assignment.teacherId && data.students.filter((s) => assignment.studentIds.includes(s.id)).some((s) => {
                          const pt = constraintFor(data.constraints, assignment.teacherId, s.id)
                          return pt === 'incompatible'
                        }) || (!isMendan && assignment.studentIds.length === 2 && constraintFor(data.constraints, assignment.studentIds[0], assignment.studentIds[1]) === 'incompatible')
                        const sig = assignmentSignature(assignment)
                        const hl = data.autoAssignHighlights ?? {}
                        const isAutoAdded = (hl.added?.[slot] ?? []).includes(sig)
                        const isAutoChanged = (hl.changed?.[slot] ?? []).includes(sig)
                        const isAutoMakeupHighlight = (hl.makeup?.[slot] ?? []).includes(sig)
                        // 振替 badge is permanent: show if regularMakeupInfo exists OR if auto-assign just detected it
                        const hasMakeupInfo = !!(assignment.regularMakeupInfo && Object.keys(assignment.regularMakeupInfo).length > 0)
                        const isAutoMakeup = hasMakeupInfo || isAutoMakeupHighlight
                        const manualConstraintWarnings = !isRecorded ? getManualConstraintWarnings(slot, assignment, idx) : []
                        const hasManualConstraintWarning = manualConstraintWarnings.length > 0
                        // Red border only for transient auto-assign highlights (not for permanent 振替)
                        const isAutoDiff = isAutoAdded || isAutoChanged || isAutoMakeupHighlight
                        const changeDetail = hl.changeDetails?.[slot]?.[sig] ?? ''

                        // Compute student-drag drop validity for this specific assignment block
                        const isStudentDropCandidate = isDragActive && isStudentDrag && !assignment.isGroupLesson && assignment.studentIds.length < 2
                        const isSameAssignment = isStudentDrag && dragInfo.sourceSlot === slot && dragInfo.sourceIdx === idx
                        const draggedStudentId = isStudentDrag ? dragInfo.studentDragId! : ''
                        const isStudentAlreadyInSlot = isStudentDrag && slotAssignments.some((a, aIdx) => {
                          if (isSameAssignment && aIdx === idx) return false // don't count source
                          if (dragInfo.sourceSlot === slot && dragInfo.sourceIdx === aIdx) return false // source being removed
                          return a.studentIds.includes(draggedStudentId)
                        })
                        const isTeacherCompatibleForDrag = (() => {
                          if (!isStudentDropCandidate || !selectedTeacher) return true
                          // For regular student: only allow their regular teacher's block
                          const regTid = dragInfo.regularTeacherId
                          if (regTid && selectedTeacher.id !== regTid) return false
                          const draggedStudent = data.students.find(s => s.id === draggedStudentId)
                          const dragSubject = isStudentDrag ? dragInfo.studentDragSubject ?? '' : ''
                          if (!draggedStudent || !dragSubject) return true
                          return canTeachSubject(selectedTeacher.subjects, draggedStudent.grade, dragSubject)
                        })()
                        const isStudentDropValid = isStudentDropCandidate && !isSameAssignment && !isStudentAlreadyInSlot && !hasUnavailableStudent && isTeacherCompatibleForDrag
                        const isStudentDropInvalid = isDragActive && isStudentDrag && !isStudentDropValid && !isSameAssignment

                        // Group lesson: show compact non-editable block with ■ marker
                        if (assignment.isGroupLesson) {
                          const tName = selectedTeacher?.name ?? '?'
                          const sNames = assignment.studentIds
                            .map((sid) => data.students.find((s) => s.id === sid)?.name ?? '?')
                          return (
                            <div key={idx} className="assignment-block assignment-block-regular"
                              style={{ position: 'relative', background: '#e0e7ff', borderColor: '#818cf8' }}>
                              <span className="badge" style={{ background: '#6366f1', color: '#fff', fontSize: '0.72em', padding: '2px 7px', marginBottom: '4px' }} title="集団授業">■</span>
                              <div style={{ fontWeight: 600, fontSize: '0.9em' }}>{tName}</div>
                              <div style={{ fontSize: '0.8em', color: '#4338ca' }}>{assignment.subject}</div>
                              <div style={{ fontSize: '0.78em', color: '#475569', marginTop: '2px' }}>
                                {sNames.map((name, i) => (
                                  <div key={i}>・{name}</div>
                                ))}
                              </div>
                            </div>
                          )
                        }

                        return (
                          <div
                            key={idx}
                            className={`assignment-block${assignment.isRegular ? ' assignment-block-regular' : ''}${isIncompatiblePair ? ' assignment-block-incompatible' : ''}${hasManualConstraintWarning ? ' assignment-block-manual-warning' : ''}${isAutoDiff ? ' assignment-block-auto-updated' : ''}${isStudentDropValid ? ' assignment-block-drop-target' : ''}${isStudentDropInvalid ? ' assignment-block-drop-invalid' : ''}`}
                            style={isDragActive && isSourceSlot && dragInfo.sourceIdx === idx ? { outline: '2px solid #3b82f6', outlineOffset: '-2px', background: isStudentDrag ? '#eff6ff' : '#fef3c7' } : undefined}
                          >
                            {/* Student-drag destination: "ここに移動" on valid target assignment */}
                            {isStudentDropValid && (
                              <button
                                className="btn"
                                type="button"
                                style={{ width: '100%', fontSize: '0.78em', padding: '3px', marginBottom: '4px', background: '#dcfce7', border: '1px solid #22c55e', color: '#15803d' }}
                                onClick={() => {
                                  void moveStudentToSlot(dragInfo.sourceSlot, dragInfo.sourceIdx, dragInfo.studentDragId!, slot, idx)
                                  setDragInfo(null)
                                  setTransferSlot(null)
                                }}
                              >
                                ここに移動
                              </button>
                            )}
                            {/* Badges row above teacher */}
                            {(isIncompatiblePair || hasManualConstraintWarning || isAutoAdded || isAutoChanged) && (
                              <div style={{ display: 'flex', gap: '4px', marginBottom: '2px', flexWrap: 'wrap' }}>
                                {isIncompatiblePair && <span className="badge incompatible-badge" title="制約不可">⚠</span>}
                                {hasManualConstraintWarning && <span className="badge manual-constraint-badge" title={manualConstraintWarnings.join('\n')}>注意</span>}
                                {isAutoAdded && !isAutoMakeup && !hasMakeupInfo && <span className="badge auto-diff-badge auto-diff-badge-new" title={changeDetail || '自動提案で新規追加'}>新規</span>}
                                {isAutoChanged && !isAutoMakeup && !hasMakeupInfo && <span className="badge auto-diff-badge auto-diff-badge-update" title={changeDetail || '自動提案で再割当'}>変更</span>}
                              </div>
                            )}
                            {/* Teacher row: select + pair move + delete */}
                            <div style={{ display: 'flex', alignItems: 'center', gap: '3px' }}>
                              <select
                                style={{ flex: 1, minWidth: 0 }}
                                value={assignment.teacherId}
                                onChange={(e) => void setSlotTeacher(slot, idx, e.target.value)}
                                disabled={assignment.isGroupLesson}
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
                              <div className="row-actions-right">
                                {!assignment.isGroupLesson && !isDragActive && (
                                  <button
                                    type="button"
                                    title="このペアを削除"
                                    style={{ background: '#e2e8f0', border: 'none', borderRadius: '50%', cursor: 'pointer', fontSize: '14px', color: '#64748b', width: '20px', height: '20px', padding: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, lineHeight: 1 }}
                                    onClick={() => void setSlotTeacher(slot, idx, '')}
                                  >
                                    ×
                                  </button>
                                )}
                                {!assignment.isGroupLesson && !isDragActive && !isRecorded && (
                                  <button
                                    type="button"
                                    title="ペアを別コマへ移動"
                                    style={{ background: '#eff6ff', border: '1px solid #93c5fd', borderRadius: '4px', cursor: 'pointer', fontSize: '0.9em', color: '#2563eb', padding: '1px 5px', lineHeight: 1, flexShrink: 0 }}
                                    onClick={() => {
                                      setDragInfo({ sourceSlot: slot, sourceIdx: idx, teacherId: assignment.teacherId, studentIds: [...assignment.studentIds] })
                                      setTransferSlot(slot)
                                    }}
                                  >
                                    ⇔
                                  </button>
                                )}
                              </div>
                            </div>

                            {assignment.teacherId && (
                              <>
                                {(isMendan ? [0] : [0, 1]).map((pos) => {
                                  const otherStudentId = assignment.studentIds[pos === 0 ? 1 : 0] ?? ''
                                  const currentStudentId = assignment.studentIds[pos] ?? ''
                                  const currentStudent = data.students.find((s) => s.id === currentStudentId)
                                  const studentSubject = currentStudentId
                                    ? getStudentSubject(assignment, currentStudentId)
                                    : ''
                                  // Manual assignment: show all subjects this teacher can handle for the student's grade.
                                  const studentSubjectOptions = selectedTeacher && currentStudent
                                    ? teachableBaseSubjects(selectedTeacher.subjects, currentStudent.grade)
                                    : (selectedTeacher ? [...new Set(selectedTeacher.subjects.map(s => getSubjectBase(s)))] : [])
                                  // Ensure currently assigned subject is always in the options (e.g. regular lessons)
                                  if (studentSubject && !studentSubjectOptions.includes(studentSubject)) {
                                    studentSubjectOptions.unshift(studentSubject)
                                  }
                                  // Compute ★ badge for this student (shown in left column)
                                  const starBadge = (() => {
                                    if (!currentStudentId || isMendan) return null
                                    const DAY_NAMES_STAR = ['日', '月', '火', '水', '木', '金', '土']
                                    const isRegAtSlot = assignment.isRegular && findRegularLessonsForSlot(data.regularLessons, slot).some(r => r.studentIds.includes(currentStudentId))
                                    const mkInfo = assignment.regularMakeupInfo?.[currentStudentId]
                                    if (isRegAtSlot) return <span className="badge regular-badge" style={{ fontSize: '0.7em', verticalAlign: 'middle' }} title="通常授業">★</span>
                                    if (mkInfo) {
                                      const fmtMkDate = (d: string) => { const [, m, day] = d.split('-'); return `${Number(m)}/${Number(day)}` }
                                      const origLabel = mkInfo.date ? `${fmtMkDate(mkInfo.date)} ${mkInfo.slotNumber}限` : `${DAY_NAMES_STAR[mkInfo.dayOfWeek]}曜${mkInfo.slotNumber}限`
                                      const [curDate] = slot.split('_')
                                      const curSlotNum = getSlotNumber(slot)
                                      const destLabel = `${fmtMkDate(curDate)} ${curSlotNum}限`
                                      return <span className="badge regular-badge" style={{ fontSize: '0.7em', verticalAlign: 'middle', background: '#16a34a' }} title={`振替（${origLabel} → ${destLabel}）`}>★</span>
                                    }
                                    return null
                                  })()
                                  const isSourceStudent = isDragActive && isStudentDrag && isSourceSlot && dragInfo.sourceIdx === idx && dragInfo.studentDragId === currentStudentId
                                  return (
                                    <div key={pos} className="student-select-row"
                                      style={isSourceStudent ? { background: '#dbeafe', borderRadius: '4px', outline: '2px solid #3b82f6', outlineOffset: '-1px' } : undefined}
                                    >
                                    <div className="star-badge-col">{starBadge}</div>
                                    <select
                                      value={currentStudentId}
                                      disabled={assignment.isGroupLesson}
                                      onChange={(e) => {
                                        const selectedId = e.target.value
                                        if (selectedId && !isMendan) {
                                          const student = data.students.find((s) => s.id === selectedId)
                                          if (student) {
                                            const pairTag = constraintFor(data.constraints, assignment.teacherId, student.id)
                                            const studentStudentTag = otherStudentId ? constraintFor(data.constraints, otherStudentId, student.id) : null
                                            const isIncompatible = pairTag === 'incompatible' || studentStudentTag === 'incompatible'
                                            if (isIncompatible) {
                                              const reasons: string[] = []
                                              if (pairTag === 'incompatible') reasons.push('講師×生徒ペア制約で不可')
                                              if (studentStudentTag === 'incompatible') reasons.push('生徒×生徒ペア制約で不可')
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
                                          return true
                                        })
                                        .map((student) => {
                                        const pairTag = isMendan ? null : constraintFor(data.constraints, assignment.teacherId, student.id)
                                        const ssTag = (!isMendan && otherStudentId) ? constraintFor(data.constraints, otherStudentId, student.id) : null
                                        const isIncompatible = pairTag === 'incompatible' || ssTag === 'incompatible'
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
                                    <div className="row-actions-right">
                                    {currentStudentId && !assignment.isGroupLesson && !isMendan && (
                                      <select
                                        className="subject-select-inline"
                                        value={studentSubject}
                                        onChange={(e) => void setSlotSubject(slot, idx, e.target.value, currentStudentId)}
                                      >
                                        {studentSubjectOptions.map((subj) => (
                                          <option key={subj} value={subj}>{currentStudent ? `${currentStudent.grade} ${subj}` : subj}</option>
                                        ))}
                                      </select>
                                    )}
                                    {currentStudentId && assignment.isGroupLesson && !isMendan && (
                                      <span className="subject-label-inline">{currentStudent ? `${currentStudent.grade} ${studentSubject}` : studentSubject}</span>
                                    )}
                                    {currentStudentId && !assignment.isGroupLesson && !isDragActive && !isRecorded && (
                                      <button
                                        type="button"
                                        title="生徒を別コマへ移動"
                                        style={{ cursor: 'pointer', background: '#eff6ff', border: '1px solid #93c5fd', borderRadius: '3px', padding: '0 4px', color: '#2563eb', fontSize: '0.8em', lineHeight: 1.4, flexShrink: 0, marginLeft: 2 }}
                                        onClick={() => {
                                          // Find the regular lesson teacher for this student (if any)
                                          const regLesson = data.regularLessons.find(r => r.studentIds.includes(currentStudentId))
                                          const isRegularStudent = !!(starBadge || regLesson)
                                          setDragInfo({
                                            sourceSlot: slot,
                                            sourceIdx: idx,
                                            teacherId: '',
                                            studentIds: [currentStudentId],
                                            studentDragId: currentStudentId,
                                            studentDragSubject: studentSubject,
                                            ...(isRegularStudent && regLesson ? { regularTeacherId: regLesson.teacherId } : {}),
                                          })
                                          setTransferSlot(slot)
                                        }}
                                      >
                                        ⇔
                                      </button>
                                    )}
                                    </div>
                                    </div>
                                  )
                                })}
                              </>
                            )}

                            {hasManualConstraintWarning && (
                              <div className="manual-constraint-note">
                                <strong>手動割当で制約カードを超過中</strong>
                                {manualConstraintWarnings.map((warning) => (
                                  <div key={warning}>{warning}</div>
                                ))}
                              </div>
                            )}

                          </div>
                        )
                      })}
                      {!isDragActive && (() => {
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
                    )}
                  </div>
                )
              })}
            </div>
            </div>{/* close loading overlay wrapper */}
          </div>
          {statusModal && (
            <div className="status-modal-backdrop" onClick={() => setStatusModal(null)}>
              <div className="status-modal" onClick={(e) => e.stopPropagation()}>
                <div className="status-modal-header">
                  <div>
                    <h3>{statusModal.title}</h3>
                    <p>{statusModal.summary}</p>
                  </div>
                  <button className="btn secondary" type="button" onClick={() => setStatusModal(null)}>閉じる</button>
                </div>
                <div className="status-modal-body">
                  {statusModal.sections.map((section) => (
                    <section key={section.key} className="status-section">
                      <h4>{section.title}</h4>
                      {section.items.map((item, idx) => (
                        <div key={`${section.key}-${idx}-${item.label}`} className="status-item-card">
                          <div className="status-item-title">{item.label}</div>
                          <div className="status-item-block">
                            <strong>原因</strong>
                            <ul>
                              {item.causes.map((cause, causeIdx) => <li key={causeIdx}>{cause}</li>)}
                            </ul>
                          </div>
                          <div className="status-item-block">
                            <strong>提案</strong>
                            <ul>
                              {(item.proposals.length > 0 ? item.proposals : [toStatusProposal('提案候補なし')]).map((proposal, proposalIdx) => (
                                <li key={proposalIdx}>
                                  {proposal.action ? (
                                    <button
                                      className="status-proposal-action"
                                      type="button"
                                      onClick={() => void applyForceAssignAction(proposal.action!)}
                                    >
                                      {proposal.label}
                                    </button>
                                  ) : proposal.label}
                                </li>
                              ))}
                            </ul>
                          </div>
                        </div>
                      ))}
                    </section>
                  ))}
                </div>
              </div>
            </div>
          )}
        </>
      )}

    </div>
  )
}

// Teacher Input Component
// Teacher/Manager Input Component
const TeacherInputPage = ({
  classroomId,
  sessionId,
  data,
  teacher,
  returnToAdminOnComplete,
  personKeyPrefix = 'teacher',
}: {
  classroomId: string
  sessionId: string
  data: SessionData
  teacher: Teacher
  returnToAdminOnComplete: boolean
  personKeyPrefix?: 'teacher' | 'manager'
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const showDevRandom = true
  const formRef = useRef<HTMLDivElement>(null)

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
    saveSession(classroomId, sessionId, next).catch(() => { /* ignore */ })
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
      captureElement: formRef.current,
    }).finally(() => {
      navigate(`/c/${classroomId}/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
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
      <div className="availability-container" ref={formRef}>
        <div className="availability-header">
          {returnToAdminOnComplete && (
            <button className="btn secondary" type="button" style={{ marginBottom: '8px', fontSize: '0.85em' }}
              onClick={() => navigate(-1)}>← 戻る</button>
          )}
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
    <div className="availability-container" ref={formRef}>
      <div className="availability-header">
        {returnToAdminOnComplete && (
          <button className="btn secondary" type="button" style={{ marginBottom: '8px', fontSize: '0.85em' }}
            onClick={() => navigate(-1)}>← 戻る</button>
        )}
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
  classroomId,
  sessionId,
  data,
  student,
  returnToAdminOnComplete,
}: {
  classroomId: string
  sessionId: string
  data: SessionData
  student: Student
  returnToAdminOnComplete: boolean
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const showDevRandom = true
  const formRef = useRef<HTMLDivElement>(null)
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

  const regularSlotMap = useMemo(() => {
    const map = new Map<string, string>()
    const studentLessons = data.regularLessons.filter((lesson) => lesson.studentIds.includes(student.id))
    for (const date of dates) {
      const dayOfWeek = getIsoDayOfWeek(date)
      for (const lesson of studentLessons) {
        if (lesson.dayOfWeek === dayOfWeek) {
          const subj = lesson.studentSubjects?.[student.id] ?? lesson.subject ?? ''
          map.set(`${date}_${lesson.slotNumber}`, subj)
        }
      }
    }
    return map
  }, [dates, data.regularLessons, student.id])
  const regularSlotKeys = useMemo(() => new Set(regularSlotMap.keys()), [regularSlotMap])

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
      unavailableDates: derivedUnavailDates,
      preferredSlots: [],
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
    saveSession(classroomId, sessionId, next).catch(() => { /* ignore */ })
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
      captureElement: formRef.current,
    }).finally(() => {
      navigate(`/c/${classroomId}/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
    })
  }

  return (    <div className="availability-container" ref={formRef}>
      <div className="availability-header">
        {returnToAdminOnComplete && (
          <button className="btn secondary" type="button" style={{ marginBottom: '8px', fontSize: '0.85em' }}
            onClick={() => navigate(-1)}>← 戻る</button>
        )}
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
                            className={`teacher-slot-btn ${isUnavail && hasRegular ? 'unavail-regular' : isUnavail ? 'unavail' : ''} ${hasRegular && !isUnavail ? 'regular' : ''}`}
                            onClick={() => toggleSlot(slotKey)}
                            type="button"
                            style={hasRegular && !isUnavail ? { fontSize: '11px', lineHeight: '1.1', padding: '2px 1px' } : undefined}
                          >
                            {isUnavail ? '✕' : hasRegular ? <>通常<br />{regularSlotMap.get(slotKey) ?? ''}</> : ''}
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
  classroomId,
  sessionId,
  data,
  student,
  returnToAdminOnComplete,
}: {
  classroomId: string
  sessionId: string
  data: SessionData
  student: Student
  returnToAdminOnComplete: boolean
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const showDevRandom = true
  const formRef = useRef<HTMLDivElement>(null)

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
    saveSession(classroomId, sessionId, next).catch(() => { /* ignore */ })
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
      captureElement: formRef.current,
    }).finally(() => {
      navigate(`/c/${classroomId}/complete/${sessionId}`, { state: { returnToAdminOnComplete } })
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
    <div className="availability-container" ref={formRef}>
      <div className="availability-header">
        {returnToAdminOnComplete && (
          <button className="btn secondary" type="button" style={{ marginBottom: '8px', fontSize: '0.85em' }}
            onClick={() => navigate(-1)}>← 戻る</button>
        )}
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
      const startSlot = isMendan ? 1 : 0  // Include slot 0 (午前) for non-mendan
      for (let s = startSlot; s <= slotsPerDay; s++) {
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
                const isGroupSlot = s === 0
                cells.push({
                  label: student?.name ?? '?',
                  detail: isGroupSlot ? `${subj} (集団)` : subj,
                  color: isGroupSlot ? '#e0e7ff' : a.isRegular ? '#dcfce7' : '#fef3c7',
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
              const isGroupSlot = s === 0
              const isRegAtSlot = a.isRegular && findRegularLessonsForSlot(data.regularLessons, `${date}_${s}`).some(r => r.studentIds.includes(personId))
              const mkInfo = a.regularMakeupInfo?.[personId]
              const isRegularStudent = isRegAtSlot || !!mkInfo
              const DAY_LABELS_PS = ['日', '月', '火', '水', '木', '金', '土']
              const makeupHint = mkInfo ? ` (${DAY_LABELS_PS[mkInfo.dayOfWeek]}曜${mkInfo.slotNumber}限の振替)` : ''
              cells.push({
                label: isGroupSlot ? `■ ${subj}` : isRegularStudent ? `★ ${subj}` : subj,
                detail: `${teacher?.name ?? '?'}${isGroupSlot ? ' (集団授業)' : isRegularStudent ? ` (通常${makeupHint})` : ' (特別講習)'}`,
                color: isGroupSlot ? '#e0e7ff' : isRegularStudent ? '#dcfce7' : '#fef3c7',
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
    if (!isMendan) return Array.from({ length: slotsPerDay + 1 }, (_, i) => i)  // 0 (午前), 1..slotsPerDay
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
    if (slotNum === 0) return '午前'
    return `${slotNum}限`
  }

  const roleLabel = personType === 'teacher' ? '講師' : personType === 'manager' ? 'マネージャー' : isMendan ? '保護者' : '生徒'

  // Count total assigned slots for this person
  const totalSlots = useMemo(() => {
    let count = 0
    for (const date of dates) {
      const startSlot = isMendan ? 1 : 0
      for (let s = startSlot; s <= slotsPerDay; s++) {
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
  const { classroomId = '', sessionId = 'main', personType: rawPersonType = 'teacher', personId: rawPersonId = '' } = useParams()
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
      classroomId,
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
        loadMasterData(classroomId)
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
              regularLessons: master.regularLessons,
            }
            setData(next)
            setPhase('ready')
            return saveSession(classroomId, sessionId, next)
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
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
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
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
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
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
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
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
        </div>
      </div>
    )
  }

  // When admin has confirmed the schedule, show calendar view instead of input form
  if (data.settings.confirmed) {
    return <ConfirmedCalendarView data={data} personType={personType} personId={personId} />
  }

  // Submission period check (proxy input from admin skips deadline)
  const now = new Date()
  const subStartDate = data.settings.submissionStartDate ? new Date(data.settings.submissionStartDate) : null
  const subEndDate = data.settings.submissionEndDate ? new Date(data.settings.submissionEndDate + 'T23:59:59') : null
  const isBeforeStart = subStartDate ? now < subStartDate : false
  const isAfterEnd = subEndDate ? now > subEndDate : false

  // Before submission period: block (but not for proxy input)
  if (isBeforeStart && !returnToAdminOnComplete) {
    return (
      <div className="app-shell">
        <div className="panel">
          <h3>提出期間前です</h3>
          <p>提出受付開始日: <strong>{data.settings.submissionStartDate}</strong></p>
          <p className="muted">提出期間になるまでお待ちください。</p>
          <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
        </div>
      </div>
    )
  }

  // After submission period: show calendar view (read-only) instead of blocking
  if (isAfterEnd && !returnToAdminOnComplete) {
    return (
      <div className="app-shell">
        <div className="panel">
          <h3>提出期間は終了しました</h3>
          <p>提出締切日: <strong>{data.settings.submissionEndDate}</strong></p>
          <p className="muted">期間を過ぎているため入力はできません。以下のスケジュールをご確認ください。</p>
        </div>
        <ConfirmedCalendarView data={data} personType={personType} personId={personId} />
      </div>
    )
  }

  if (personType === 'teacher') {
    if ('subjects' in currentPerson && Array.isArray(currentPerson.subjects)) {
      return <TeacherInputPage classroomId={classroomId} sessionId={sessionId} data={data} teacher={currentPerson as Teacher} returnToAdminOnComplete={returnToAdminOnComplete} />
    }
  } else if (personType === 'manager') {
    // Manager availability input — same as teacher but uses 'manager' personKey
    const manager = currentPerson as Manager
    // Wrap manager as a Teacher-like object so TeacherInputPage can be reused
    const managerAsTeacher: Teacher = { id: manager.id, name: manager.name, email: manager.email, subjects: ['面談'], memo: '' }
    return <TeacherInputPage classroomId={classroomId} sessionId={sessionId} data={data} teacher={managerAsTeacher} returnToAdminOnComplete={returnToAdminOnComplete} personKeyPrefix="manager" />
  } else if (personType === 'student') {
    if ('grade' in currentPerson && 'subjectSlots' in currentPerson) {
      return data.settings.sessionType === 'mendan'
        ? <MendanParentInputPage classroomId={classroomId} sessionId={sessionId} data={data} student={currentPerson as Student} returnToAdminOnComplete={returnToAdminOnComplete} />
        : <StudentInputPage classroomId={classroomId} sessionId={sessionId} data={data} student={currentPerson as Student} returnToAdminOnComplete={returnToAdminOnComplete} />
    }
  }

  return (
    <div className="app-shell">
      <div className="panel">
        入力対象の種別が正しくありません。管理者にURLを確認してください。
        <br />
        <Link to={`/c/${classroomId}`}>ホームに戻る</Link>
      </div>
    </div>
  )
}
/** Legacy redirect: old URLs without classroomId (e.g. /availability/:sessionId/:personType/:personId) */
const LegacyAvailabilityRedirect = () => {
  const navigate = useNavigate()
  const { sessionId = '', personType = '', personId = '' } = useParams()
  const [error, setError] = useState('')

  useEffect(() => {
    let cancelled = false
    initAuth().then(() => findClassroomForSession(sessionId)).then((classroomId) => {
      if (cancelled) return
      if (classroomId) {
        navigate(`/c/${classroomId}/availability/${sessionId}/${personType}/${personId}`, { replace: true })
      } else {
        setError('該当する教室が見つかりません。URLが正しいかご確認ください。')
      }
    }).catch(() => {
      if (!cancelled) setError('データの読み込みに失敗しました。')
    })
    return () => { cancelled = true }
  }, [sessionId, personType, personId, navigate])

  if (error) {
    return (
      <div className="app-shell">
        <div className="panel">
          <h3>エラー</h3>
          <p>{error}</p>
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <div className="panel">
        <p>接続中…</p>
      </div>
    </div>
  )
}

const ClassroomSelectPage = () => {
  const navigate = useNavigate()
  const [classrooms, setClassrooms] = useState<ClassroomInfo[]>([])
  const [newId, setNewId] = useState('')
  const [newName, setNewName] = useState('')
  const [loading, setLoading] = useState(true)
  const [backupClassroomId, setBackupClassroomId] = useState<string | null>(null)
  const [backups, setBackups] = useState<BackupMeta[]>([])
  const [backupsLoading, setBackupsLoading] = useState(false)
  const [backupBusy, setBackupBusy] = useState(false)

  useEffect(() => {
    const unsub = watchClassrooms((items) => {
      setClassrooms(items)
      setLoading(false)
    })
    return () => unsub()
  }, [])

  const handleCreate = async (): Promise<void> => {
    const id = newId.trim().toLowerCase().replace(/[^a-z0-9_-]/g, '')
    const name = newName.trim()
    if (!id || !name) { alert('IDと名前を入力してください。'); return }
    if (classrooms.some((c) => c.id === id)) { alert('同じIDの教室が既に存在します。'); return }
    await createClassroom(id, name)
    setNewId('')
    setNewName('')
  }

  const handleDelete = async (id: string, name: string): Promise<void> => {
    if (!window.confirm(`教室「${name}」を削除しますか？\nこの操作は元に戻せません。すべてのセッションデータも削除されます。`)) return
    await deleteClassroom(id)
  }

  const refreshBackups = useCallback(async (cId: string) => {
    setBackupsLoading(true)
    try {
      const items = await listBackups(cId)
      setBackups(items)
    } catch (e) {
      console.error('Failed to load backups:', e)
      setBackups([])
    }
    setBackupsLoading(false)
  }, [])

  const openBackupPanel = async (cId: string): Promise<void> => {
    if (backupClassroomId === cId) {
      setBackupClassroomId(null)
      return
    }
    setBackupClassroomId(cId)
    await refreshBackups(cId)
  }

  const handleManualBackup = async (cId: string): Promise<void> => {
    setBackupBusy(true)
    try {
      await createBackup(cId, 'manual')
      await cleanupOldBackups(cId, 30)
      await refreshBackups(cId)
      alert('バックアップを作成しました。')
    } catch (e) {
      alert(`バックアップ作成エラー: ${e instanceof Error ? e.message : String(e)}`)
    }
    setBackupBusy(false)
  }

  const handleRestore = async (cId: string, backupId: string, createdAt: number): Promise<void> => {
    const dateStr = new Date(createdAt).toLocaleString('ja-JP')
    if (!window.confirm(`${dateStr} のバックアップに復元しますか？\n\n現在のデータは上書きされます。この操作は元に戻せません。`)) return
    setBackupBusy(true)
    try {
      const backup = await loadBackup(cId, backupId)
      if (!backup) { alert('バックアップデータの読み込みに失敗しました。'); setBackupBusy(false); return }
      await restoreBackup(cId, backup)
      alert('復元が完了しました。')
    } catch (e) {
      alert(`復元エラー: ${e instanceof Error ? e.message : String(e)}`)
    }
    setBackupBusy(false)
  }

  const handleDeleteBackup = async (cId: string, backupId: string): Promise<void> => {
    if (!window.confirm('このバックアップを削除しますか？')) return
    try {
      await deleteBackup(cId, backupId)
      await refreshBackups(cId)
    } catch (e) {
      alert(`削除エラー: ${e instanceof Error ? e.message : String(e)}`)
    }
  }

  const formatBackupDate = (ts: number): string => new Date(ts).toLocaleString('ja-JP')

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>教室選択</h2>
        <p className="muted">管理する教室を選択してください。</p>
      </div>

      {loading ? (
        <div className="panel">読み込み中...</div>
      ) : (
        <>
          {classrooms.length === 0 ? (
            <div className="panel">
              <p>教室がまだ登録されていません。下のフォームから教室を作成してください。</p>
            </div>
          ) : (
            <div className="panel">
              <table className="table">
                <thead><tr><th>教室名</th><th>ID</th><th></th><th></th><th></th></tr></thead>
                <tbody>
                  {classrooms.map((c) => (
                    <tr key={c.id}>
                      <td><strong>{c.name}</strong></td>
                      <td className="muted">{c.id}</td>
                      <td><button className="btn" type="button" onClick={() => navigate(`/c/${c.id}`)}>開く</button></td>
                      <td><button className="btn secondary" type="button" onClick={() => void openBackupPanel(c.id)}>{backupClassroomId === c.id ? '閉じる' : '🔄 データ復元'}</button></td>
                      <td><button className="btn secondary" type="button" style={{ color: '#dc2626' }} onClick={() => void handleDelete(c.id, c.name)}>削除</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {backupClassroomId && (
            <div className="panel">
              <h3>データ復元 — {classrooms.find((c) => c.id === backupClassroomId)?.name ?? backupClassroomId}</h3>
              <p className="muted">「保存して閉じる」ボタンを押すたびにバックアップが作成されます（最大30件保持）。復元すると、その時点の講師・生徒・制約・通常授業・全セッションのデータが復元されます。</p>
              <div className="row" style={{ marginBottom: '8px', gap: '8px' }}>
                <button className="btn" type="button" disabled={backupBusy} onClick={() => void handleManualBackup(backupClassroomId)}>
                  {backupBusy ? '処理中...' : '📸 今すぐバックアップ'}
                </button>
              </div>
              {backupsLoading ? (
                <p>読み込み中...</p>
              ) : backups.length === 0 ? (
                <p className="muted">バックアップはまだありません。</p>
              ) : (
                <table className="table">
                  <thead><tr><th>日時</th><th>種別</th><th>変更内容</th><th>操作</th></tr></thead>
                  <tbody>
                    {backups.map((b) => (
                      <tr key={b.id}>
                        <td>{formatBackupDate(b.createdAt)}</td>
                        <td>{b.trigger === 'auto' ? '🤖 自動' : '👤 手動'}</td>
                        <td>
                          {b.changeLog.length > 0
                            ? b.changeLog.join('、')
                            : `講師${b.teacherCount}名・生徒${b.studentCount}名${b.sessionCount > 0 ? ` / ${b.sessionNames.join(', ')}` : ''}`
                          }
                        </td>
                        <td>
                          <button className="btn secondary" type="button" style={{ marginRight: '4px' }} disabled={backupBusy} onClick={() => void handleRestore(backupClassroomId, b.id, b.createdAt)}>復元</button>
                          <button className="btn secondary" type="button" style={{ color: '#dc2626' }} onClick={() => void handleDeleteBackup(backupClassroomId, b.id)}>削除</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}

          <div className="panel">
            <h3>新しい教室を作成</h3>
            <div className="row" style={{ gap: '8px', flexWrap: 'wrap' }}>
              <input type="text" placeholder="教室ID（英数字）" value={newId} onChange={(e) => setNewId(e.target.value)} style={{ width: '150px' }} />
              <input type="text" placeholder="教室名" value={newName} onChange={(e) => setNewName(e.target.value)} style={{ width: '200px' }} />
              <button className="btn" type="button" onClick={() => void handleCreate()}>作成</button>
            </div>
          </div>


        </>
      )}
    </div>
  )
}

const BootPage = () => {
  const navigate = useNavigate()
  const classroomId = 'default'

  useEffect(() => {
    saveSession(classroomId, 'main', createTemplateSession()).catch(() => {})
    navigate(`/c/${classroomId}/admin/main`, { replace: true })
  }, [navigate])

  return (
    <div className="app-shell">
      <div className="panel">初期化中...</div>
    </div>
  )
}

const CompletionPage = () => {
  const location = useLocation()
  const { classroomId = '', sessionId = 'main' } = useParams()
  const returnToAdminOnComplete = (location.state as { returnToAdminOnComplete?: boolean } | null)?.returnToAdminOnComplete === true

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>入力完了</h2>
        <p>データの送信が完了しました。ありがとうございます。</p>
        {returnToAdminOnComplete && (
          <div className="row" style={{ marginTop: '10px' }}>
            <Link className="btn" to={`/c/${classroomId}/admin/${sessionId}`} state={{ skipAuth: true }}>管理画面へ戻る</Link>
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
        <Route path="/" element={<ClassroomSelectPage />} />
        <Route path="/c/:classroomId" element={<HomePage />} />
        <Route path="/c/:classroomId/boot" element={<BootPage />} />
        <Route path="/c/:classroomId/admin/:sessionId" element={<AdminPage />} />
        <Route path="/c/:classroomId/availability/:sessionId/:personType/:personId" element={<AvailabilityPage />} />
        <Route path="/c/:classroomId/complete/:sessionId" element={<CompletionPage />} />
        {/* Legacy route: old URLs without classroomId */}
        <Route path="/availability/:sessionId/:personType/:personId" element={<LegacyAvailabilityRedirect />} />
      </Routes>
    </>
  )
}

export default App
