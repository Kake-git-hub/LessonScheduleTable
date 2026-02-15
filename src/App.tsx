import { useEffect, useMemo, useState, type ChangeEvent } from 'react'
import { Link, Route, Routes, useLocation, useNavigate, useParams } from 'react-router-dom'
import './App.css'
import { loadSession, saveSession, watchSession, watchSessionsList } from './firebase'
import type {
  Assignment,
  ConstraintType,
  PairConstraint,
  PersonType,
  RegularLesson,
  SessionData,
  Student,
  Teacher,
} from './types'
import { buildSlotKeys, personKey, slotLabel } from './utils/schedule'

const APP_VERSION = '0.1.0'

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
  },
  subjects: ['数学', '英語'],
  teachers: [],
  students: [],
  constraints: [],
  availability: {},
  assignments: {},
  regularLessons: [],
})

const createTemplateSession = (): SessionData => {
  const settings: SessionData['settings'] = {
    name: '夏期講習テンプレート',
    adminPassword: 'admin1234',
    startDate: '2026-07-21',
    endDate: '2026-07-23',
    slotsPerDay: 3,
    holidays: [],
  }

  const subjects = ['数学', '英語']

  const teachers: Teacher[] = [
    { id: 't001', name: '田中先生', subjects: ['数学', '英語'], memo: '数学メイン' },
    { id: 't002', name: '佐藤先生', subjects: ['英語', '数学'], memo: '英語メイン' },
  ]

  const students: Student[] = [
    {
      id: 's001',
      name: '青木 太郎',
      grade: '中3',
      subjects: ['数学', '英語'],
      subjectSlots: { 数学: 3, 英語: 2 },
      unavailableDates: ['2026-07-23'],
      memo: '受験対策',
      submittedAt: Date.now() - 4000,
    },
    {
      id: 's002',
      name: '伊藤 花',
      grade: '中2',
      subjects: ['英語'],
      subjectSlots: { 英語: 3 },
      unavailableDates: [],
      memo: '',
      submittedAt: Date.now() - 3000,
    },
    {
      id: 's003',
      name: '上田 陽介',
      grade: '高1',
      subjects: ['数学'],
      subjectSlots: { 数学: 3 },
      unavailableDates: ['2026-07-22'],
      memo: '',
      submittedAt: Date.now() - 2000,
    },
    {
      id: 's004',
      name: '岡本 美咲',
      grade: '高2',
      subjects: ['英語', '数学'],
      subjectSlots: { 英語: 2, 数学: 2 },
      unavailableDates: [],
      memo: '',
      submittedAt: Date.now() - 1000,
    },
  ]

  const constraints: PairConstraint[] = [
    { id: 'c001', teacherId: 't001', studentId: 's002', type: 'incompatible' },
    { id: 'c002', teacherId: 't001', studentId: 's003', type: 'recommended' },
    { id: 'c003', teacherId: 't002', studentId: 's004', type: 'recommended' },
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
      subject: '数学',
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
    availability,
    assignments: {},
    regularLessons,
  }
}

const useSessionData = (sessionId: string) => {
  const [data, setData] = useState<SessionData | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    setLoading(true)
    const unsub = watchSession(sessionId, (value) => {
      setData(value)
      setLoading(false)
    })
    return () => unsub()
  }, [sessionId])

  return { data, setData, loading }
}

const constraintFor = (
  constraints: PairConstraint[],
  teacherId: string,
  studentId: string,
): ConstraintType | null => {
  const hit = constraints.find((item) => item.teacherId === teacherId && item.studentId === studentId)
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

const countStudentLoad = (assignments: Record<string, Assignment[]>, studentId: string): number =>
  allAssignments(assignments).filter((a) => a.studentIds.includes(studentId)).length

const countStudentSubjectLoad = (
  assignments: Record<string, Assignment[]>,
  studentId: string,
  subject: string,
): number =>
  allAssignments(assignments).filter(
    (a) => a.studentIds.includes(studentId) && a.subject === subject,
  ).length

const getTotalRemainingSlots = (
  assignments: Record<string, Assignment[]>,
  student: Student,
): number =>
  Object.entries(student.subjectSlots).reduce((sum, [subject, desired]) => {
    const assigned = countStudentSubjectLoad(assignments, student.id, subject)
    return sum + Math.max(0, desired - assigned)
  }, 0)

const isStudentAvailable = (student: Student, slotKey: string): boolean => {
  const [date] = slotKey.split('_')
  return !student.unavailableDates.includes(date)
}

const getSlotDayOfWeek = (slotKey: string): number => {
  const [date] = slotKey.split('_')
  const [year, month, day] = date.split('-').map(Number)
  return new Date(year, month - 1, day).getDay()
}

const getSlotNumber = (slotKey: string): number => {
  const [, slot] = slotKey.split('_')
  return Number.parseInt(slot, 10)
}

const ADMIN_PASSWORD_STORAGE_KEY = 'lst_admin_password_v1'
const readSavedAdminPassword = (): string => localStorage.getItem(ADMIN_PASSWORD_STORAGE_KEY) ?? 'admin1234'
const saveAdminPassword = (password: string): void => {
  localStorage.setItem(ADMIN_PASSWORD_STORAGE_KEY, password)
}

const findRegularLessonForSlot = (
  regularLessons: RegularLesson[],
  slotKey: string,
): RegularLesson | null => {
  const dayOfWeek = getSlotDayOfWeek(slotKey)
  const slotNumber = getSlotNumber(slotKey)
  return (
    regularLessons.find((lesson) => lesson.dayOfWeek === dayOfWeek && lesson.slotNumber === slotNumber) ?? null
  )
}

const buildAutoAssignments = (
  data: SessionData,
  slots: string[],
  onlyEmpty: boolean,
): Record<string, Assignment[]> => {
  const nextAssignments: Record<string, Assignment[]> = onlyEmpty
    ? Object.fromEntries(Object.entries(data.assignments).map(([k, v]) => [k, [...v]]))
    : {}

  for (const slot of slots) {
    if (onlyEmpty && nextAssignments[slot] && nextAssignments[slot].length > 0) {
      continue
    }

    // Rule 2: Check if there's a regular lesson for this slot
    const regularLesson = findRegularLessonForSlot(data.regularLessons, slot)
    if (regularLesson) {
      nextAssignments[slot] = [
        {
          teacherId: regularLesson.teacherId,
          studentIds: regularLesson.studentIds,
          subject: regularLesson.subject,
        },
      ]
      continue
    }

    const slotAssignments: Assignment[] = []
    const usedTeacherIds = new Set<string>()
    const usedStudentIds = new Set<string>()

    const teachers = data.teachers.filter((teacher) =>
      hasAvailability(data.availability, 'teacher', teacher.id, slot),
    )

    for (const teacher of teachers) {
      if (usedTeacherIds.has(teacher.id)) continue

      const candidates = data.students.filter((student) => {
        if (usedStudentIds.has(student.id)) return false
        if (!hasAvailability(data.availability, 'student', student.id, slot)) return false
        if (!isStudentAvailable(student, slot)) return false
        if (constraintFor(data.constraints, teacher.id, student.id) === 'incompatible') return false
        return teacher.subjects.some((subject) => student.subjects.includes(subject))
      })

      if (candidates.length === 0) continue

      let bestPlan: { score: number; assignment: Assignment } | null = null
      const teacherLoad = countTeacherLoad(nextAssignments, teacher.id)
      const oneStudentCombos = candidates.map((student) => [student])
      const twoStudentCombos = candidates.flatMap((left, index) =>
        candidates.slice(index + 1).map((right) => [left, right]),
      )
      const allCombos = [...oneStudentCombos, ...twoStudentCombos]

      for (const combo of allCombos) {
        const commonSubjects = teacher.subjects.filter((subject) =>
          combo.every((student) => student.subjects.includes(subject)),
        )

        if (commonSubjects.length === 0) continue

        const viableSubjects = commonSubjects.filter((subject) =>
          combo.every((student) => {
            const requested = student.subjectSlots[subject] ?? 0
            const allocated = countStudentSubjectLoad(nextAssignments, student.id, subject)
            return allocated < requested
          }),
        )

        if (viableSubjects.length === 0) continue

        const recommendScore = combo.reduce((score, student) => {
          return score + (constraintFor(data.constraints, teacher.id, student.id) === 'recommended' ? 30 : 0)
        }, 0)

        const studentLoadPenalty = combo.reduce((score, student) => {
          return score + countStudentLoad(nextAssignments, student.id) * 8
        }, 0)

        const priorityBonus = combo.reduce((bonus, student) => {
          const studentIndex = data.students.findIndex((s) => s.id === student.id)
          return bonus + Math.max(0, 20 - studentIndex * 2)
        }, 0)

        const unfulfillmentBonus = combo.reduce((bonus, student) => {
          const totalRequested = Object.values(student.subjectSlots).reduce((sum, count) => sum + count, 0)
          const totalAllocated = countStudentLoad(nextAssignments, student.id)
          const fulfillmentRate = totalRequested > 0 ? totalAllocated / totalRequested : 1
          return bonus + Math.max(0, Math.floor((1 - fulfillmentRate) * 10))
        }, 0)

        const groupBonus = combo.length === 2 ? 5 : 0
        const score =
          100 +
          recommendScore +
          groupBonus +
          priorityBonus +
          unfulfillmentBonus -
          teacherLoad * 6 -
          studentLoadPenalty

        if (!bestPlan || score > bestPlan.score) {
          bestPlan = {
            score,
            assignment: {
              teacherId: teacher.id,
              studentIds: combo.map((student) => student.id),
              subject: viableSubjects[0],
            },
          }
        }
      }

      if (bestPlan) {
        slotAssignments.push(bestPlan.assignment)
        usedTeacherIds.add(teacher.id)
        for (const sid of bestPlan.assignment.studentIds) {
          usedStudentIds.add(sid)
        }
      }
    }

    if (slotAssignments.length > 0) {
      nextAssignments[slot] = slotAssignments
    }
  }

  return nextAssignments
}

const HomePage = () => {
  const navigate = useNavigate()
  const [unlocked, setUnlocked] = useState(false)
  const [adminPassword, setAdminPassword] = useState(readSavedAdminPassword())
  const [sessions, setSessions] = useState<{ id: string; name: string; createdAt: number; updatedAt: number }[]>([])
  const [newYear, setNewYear] = useState(String(new Date().getFullYear()))
  const [newTerm, setNewTerm] = useState<'spring' | 'summer' | 'winter'>('summer')
  const [newSessionId, setNewSessionId] = useState('')
  const [newSessionName, setNewSessionName] = useState('')

  useEffect(() => {
    if (import.meta.env.DEV) {
      navigate('/boot', { replace: true })
    }
  }, [navigate])

  useEffect(() => {
    const year = Number.parseInt(newYear, 10)
    const safeYear = Number.isNaN(year) ? new Date().getFullYear() : year
    const label = newTerm === 'spring' ? '春期講習' : newTerm === 'summer' ? '夏期講習' : '冬期講習'
    const idTerm = newTerm === 'spring' ? 'spring' : newTerm === 'summer' ? 'summer' : 'winter'
    setNewSessionId(`${safeYear}-${idTerm}`)
    setNewSessionName(`${safeYear} ${label}`)
  }, [newTerm, newYear])

  useEffect(() => {
    if (!unlocked) {
      return
    }

    const unsub = watchSessionsList((items) => setSessions(items))
    return () => unsub()
  }, [unlocked])

  const ensureDevSession = async (): Promise<void> => {
    const id = 'dev'
    const existing = await loadSession(id)
    if (existing) {
      return
    }
    const seed = createTemplateSession()
    seed.settings.name = '開発用セッション'
    seed.settings.adminPassword = adminPassword
    await saveSession(id, seed)
  }

  const onUnlock = async (): Promise<void> => {
    saveAdminPassword(adminPassword)
    setUnlocked(true)
    await ensureDevSession()
  }

  const onCreateSession = async (): Promise<void> => {
    const id = newSessionId.trim()
    if (!id) {
      return
    }
    if (sessions.some((s) => s.id === id)) {
      alert('同じセッションIDが既に存在します。別のIDにしてください。')
      return
    }

    const seed = emptySession()
    seed.settings.name = newSessionName.trim() || id
    seed.settings.adminPassword = adminPassword
    await saveSession(id, seed)
  }

  const openAdmin = (sessionId: string): void => {
    saveAdminPassword(adminPassword)
    navigate(`/admin/${sessionId}`)
  }

  const formatDate = (ms: number): string => {
    if (!ms) {
      return '-'
    }
    const d = new Date(ms)
    const y = d.getFullYear()
    const m = String(d.getMonth() + 1).padStart(2, '0')
    const day = String(d.getDate()).padStart(2, '0')
    const hh = String(d.getHours()).padStart(2, '0')
    const mm = String(d.getMinutes()).padStart(2, '0')
    return `${y}-${m}-${day} ${hh}:${mm}`
  }

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>講習コマ割りアプリ</h2>
        <p className="muted">セッション（講習ごと）にデータを分けて管理します。</p>

        {!unlocked ? (
          <>
            <h3>管理者パスワード</h3>
            <div className="row">
              <input
                type="password"
                value={adminPassword}
                onChange={(e) => setAdminPassword(e.target.value)}
                placeholder="管理者パスワード"
              />
              <button className="btn" type="button" onClick={() => void onUnlock()}>
                続行
              </button>
            </div>
            <p className="muted">現在は初期値を保存済みのため、入力不要で続行できます。</p>
          </>
        ) : (
          <>
            <div className="panel">
              <h3>新規セッション追加</h3>
              <div className="row">
                <input
                  value={newYear}
                  onChange={(e) => setNewYear(e.target.value)}
                  placeholder="西暦"
                  style={{ width: 120 }}
                />
                <select value={newTerm} onChange={(e) => setNewTerm(e.target.value as typeof newTerm)}>
                  <option value="spring">春期講習</option>
                  <option value="summer">夏期講習</option>
                  <option value="winter">冬期講習</option>
                </select>
                <input
                  value={newSessionId}
                  onChange={(e) => setNewSessionId(e.target.value)}
                  placeholder="sessionId (例: 2026-summer)"
                />
                <input
                  value={newSessionName}
                  onChange={(e) => setNewSessionName(e.target.value)}
                  placeholder="表示名 (例: 2026 夏期講習)"
                />
                <button className="btn" type="button" onClick={() => void onCreateSession()}>
                  追加
                </button>
              </div>
            </div>

            <div className="panel">
              <div className="row" style={{ justifyContent: 'space-between' }}>
                <h3>セッション一覧（新しい順）</h3>
                <button className="btn secondary" type="button" onClick={() => setUnlocked(false)}>
                  ロック
                </button>
              </div>
              <div className="row" style={{ marginBottom: 8 }}>
                <input
                  type="password"
                  value={adminPassword}
                  onChange={(e) => setAdminPassword(e.target.value)}
                  placeholder="管理者パスワード（共通運用を想定）"
                />
                <span className="muted">このパスワードでAdminに入ります</span>
              </div>
              <table className="table">
                <thead>
                  <tr>
                    <th>セッションID</th>
                    <th>名称</th>
                    <th>作成</th>
                    <th>更新</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {sessions.map((s) => (
                    <tr key={s.id}>
                      <td>{s.id}</td>
                      <td>{s.name}</td>
                      <td>{formatDate(s.createdAt)}</td>
                      <td>{formatDate(s.updatedAt)}</td>
                      <td>
                        <button className="btn" type="button" onClick={() => openAdmin(s.id)}>
                          管理
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <p className="muted">※先生・生徒の希望入力URLは各セッションのAdmin画面から配布します。</p>
            </div>
          </>
        )}
      </div>
    </div>
  )
}

const AdminPage = () => {
  const { sessionId = 'main' } = useParams()
  const location = useLocation()
  const skipAuth = (location.state as { skipAuth?: boolean } | null)?.skipAuth === true
  const { data, setData, loading } = useSessionData(sessionId)
  const [authorized, setAuthorized] = useState(import.meta.env.DEV || skipAuth)

  const [subjectInput, setSubjectInput] = useState('')
  const [teacherName, setTeacherName] = useState('')
  const [teacherSubjectsText, setTeacherSubjectsText] = useState('')
  const [teacherMemo, setTeacherMemo] = useState('')

  const [studentName, setStudentName] = useState('')
  const [studentGrade, setStudentGrade] = useState('')
  const [studentSubjects, setStudentSubjects] = useState<string[]>([])
  const [studentSubjectSlotsText, setStudentSubjectSlotsText] = useState('')
  const [studentUnavailableDatesText, setStudentUnavailableDatesText] = useState('')
  const [studentMemo, setStudentMemo] = useState('')

  const [constraintTeacherId, setConstraintTeacherId] = useState('')
  const [constraintStudentId, setConstraintStudentId] = useState('')
  const [constraintType, setConstraintType] = useState<ConstraintType>('incompatible')

  const [regularTeacherId, setRegularTeacherId] = useState('')
  const [regularStudentIds, setRegularStudentIds] = useState<string[]>([])
  const [regularSubject, setRegularSubject] = useState('')
  const [regularDayOfWeek, setRegularDayOfWeek] = useState('')
  const [regularSlotNumber, setRegularSlotNumber] = useState('')

  useEffect(() => {
    setAuthorized(import.meta.env.DEV || skipAuth)
  }, [sessionId, skipAuth])

  const slotKeys = useMemo(() => (data ? buildSlotKeys(data.settings) : []), [data])

  const persist = async (next: SessionData): Promise<void> => {
    setData(next)
    await saveSession(sessionId, next)
  }

  const update = async (updater: (current: SessionData) => SessionData): Promise<void> => {
    if (!data) {
      return
    }
    await persist(updater(data))
  }

  const createSession = async (): Promise<void> => {
    const seed = emptySession()
    await saveSession(sessionId, seed)
  }

  const createTemplateSessionDoc = async (): Promise<void> => {
    const seed = createTemplateSession()
    await saveSession(sessionId, seed)
  }

  useEffect(() => {
    if (!data) {
      return
    }
    if (import.meta.env.DEV || skipAuth) {
      setAuthorized(true)
      return
    }
    const password = readSavedAdminPassword()
    setAuthorized(password === data.settings.adminPassword)
  }, [data, skipAuth])

  const parseList = (text: string): string[] =>
    text
      .split(/[、,\n]/)
      .map((item) => item.trim())
      .filter(Boolean)

  const addTeacher = async (): Promise<void> => {
    if (!teacherName.trim()) {
      return
    }
    const teacher: Teacher = {
      id: createId(),
      name: teacherName.trim(),
      subjects: parseList(teacherSubjectsText),
      memo: teacherMemo.trim(),
    }

    await update((current) => ({ ...current, teachers: [...current.teachers, teacher] }))
    setTeacherName('')
    setTeacherSubjectsText('')
    setTeacherMemo('')
  }

  const addStudent = async (): Promise<void> => {
    if (!studentName.trim()) {
      return
    }

    // Parse subject slots (e.g., "数学:5, 英語:3")
    const subjectSlots: Record<string, number> = {}
    if (studentSubjectSlotsText.trim()) {
      const pairs = studentSubjectSlotsText.split(/[、,]/).map((item) => item.trim())
      for (const pair of pairs) {
        const [subject, count] = pair.split(':').map((s) => s.trim())
        if (subject && count) {
          const num = Number.parseInt(count, 10)
          if (!Number.isNaN(num) && num > 0) {
            subjectSlots[subject] = num
          }
        }
      }
    }

    // Parse unavailable dates (e.g., "2026-07-25, 2026-07-28" or line-separated)
    const unavailableDates = studentUnavailableDatesText
      .split(/[、,\n]/)
      .map((item) => item.trim())
      .filter((date) => /^\d{4}-\d{2}-\d{2}$/.test(date))

    const student: Student = {
      id: createId(),
      name: studentName.trim(),
      grade: studentGrade.trim(),
      subjects: studentSubjects,
      subjectSlots,
      unavailableDates,
      memo: studentMemo.trim(),
      submittedAt: Date.now(),
    }

    await update((current) => ({ ...current, students: [...current.students, student] }))
    setStudentName('')
    setStudentGrade('')
    setStudentSubjects([])
    setStudentSubjectSlotsText('')
    setStudentUnavailableDatesText('')
    setStudentMemo('')
  }

  const upsertConstraint = async (): Promise<void> => {
    if (!constraintTeacherId || !constraintStudentId || !data) {
      return
    }

    const newConstraint: PairConstraint = {
      id: createId(),
      teacherId: constraintTeacherId,
      studentId: constraintStudentId,
      type: constraintType,
    }

    await update((current) => {
      const filtered = current.constraints.filter(
        (item) => !(item.teacherId === constraintTeacherId && item.studentId === constraintStudentId),
      )
      return { ...current, constraints: [...filtered, newConstraint] }
    })
  }

  const addRegularLesson = async (): Promise<void> => {
    if (
      !regularTeacherId ||
      regularStudentIds.length === 0 ||
      !regularSubject ||
      !regularDayOfWeek ||
      !regularSlotNumber
    ) {
      return
    }

    const newLesson: RegularLesson = {
      id: createId(),
      teacherId: regularTeacherId,
      studentIds: regularStudentIds,
      subject: regularSubject,
      dayOfWeek: Number.parseInt(regularDayOfWeek, 10),
      slotNumber: Number.parseInt(regularSlotNumber, 10),
    }

    await update((current) => ({ ...current, regularLessons: [...current.regularLessons, newLesson] }))
    setRegularTeacherId('')
    setRegularStudentIds([])
    setRegularSubject('')
    setRegularDayOfWeek('')
    setRegularSlotNumber('')
  }

  const removeRegularLesson = async (lessonId: string): Promise<void> => {
    await update((current) => ({
      ...current,
      regularLessons: current.regularLessons.filter((lesson) => lesson.id !== lessonId),
    }))
  }

  const applyAutoAssign = async (): Promise<void> => {
    if (!data) {
      return
    }

    const nextAssignments = buildAutoAssignments(data, slotKeys, true)

    await update((current) => ({ ...current, assignments: nextAssignments }))
  }

  const applyTemplateAndAutoAssign = async (): Promise<void> => {
    const template = createTemplateSession()
    const templateSlots = buildSlotKeys(template.settings)
    const assignments = buildAutoAssignments(template, templateSlots, false)
    const next: SessionData = {
      ...template,
      assignments,
    }
    await persist(next)
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

  if (loading) {
    return (
      <div className="app-shell">
        <div className="panel">読み込み中...</div>
      </div>
    )
  }

  if (!data) {
    return (
      <div className="app-shell">
        <div className="panel">
          <h2>セッション: {sessionId}</h2>
          <div className="row">
            <button className="btn" type="button" onClick={createSession}>
              空のセッションを作成
            </button>
            <button className="btn secondary" type="button" onClick={createTemplateSessionDoc}>
              テンプレートで作成
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
          <Link to="/">ホーム</Link>
        </div>
        <p className="muted">管理者のみ編集できます。希望入力は個別URLで配布してください。</p>
      </div>

      {!authorized ? (
        <div className="panel">
          <h3>管理者パスワードが一致しません</h3>
          <p className="muted">
            トップ画面で管理者パスワードを入力し「続行」してから、もう一度このセッションを開いてください。
          </p>
          <Link to="/">トップへ戻る</Link>
        </div>
      ) : (
        <>
          <div className="panel">
            <h3>講習設定</h3>
            <div className="row">
              <button className="btn secondary" type="button" onClick={createTemplateSessionDoc}>
                初期テンプレートを再投入
              </button>
              <span className="muted">現在のデータをテンプレートで上書きします。</span>
            </div>
            <div className="row">
              <input
                value={data.settings.name}
                onChange={(e) => {
                  void update((current) => ({
                    ...current,
                    settings: { ...current.settings, name: e.target.value },
                  }))
                }}
                placeholder="講習名"
              />
              <input
                type="password"
                value={data.settings.adminPassword}
                onChange={(e) => {
                  void update((current) => ({
                    ...current,
                    settings: { ...current.settings, adminPassword: e.target.value },
                  }))
                }}
                placeholder="管理者パスワード"
              />
              <input
                type="date"
                value={data.settings.startDate}
                onChange={(e) => {
                  void update((current) => ({
                    ...current,
                    settings: { ...current.settings, startDate: e.target.value },
                  }))
                }}
              />
              <input
                type="date"
                value={data.settings.endDate}
                onChange={(e) => {
                  void update((current) => ({
                    ...current,
                    settings: { ...current.settings, endDate: e.target.value },
                  }))
                }}
              />
              <input
                type="number"
                min={1}
                max={10}
                value={data.settings.slotsPerDay}
                onChange={(e) => {
                  const parsed = Number(e.target.value)
                  if (!Number.isNaN(parsed)) {
                    void update((current) => ({
                      ...current,
                      settings: { ...current.settings, slotsPerDay: parsed },
                    }))
                  }
                }}
              />
            </div>
            <div className="list">
              <div className="muted">休日(YYYY-MM-DDを改行orカンマ区切り)</div>
              <textarea
                value={data.settings.holidays.join('\n')}
                onChange={(e) => {
                  const holidays = parseList(e.target.value)
                  void update((current) => ({
                    ...current,
                    settings: { ...current.settings, holidays },
                  }))
                }}
              />
            </div>
          </div>

          <div className="panel">
            <h3>科目マスター</h3>
            <div className="row">
              <input
                value={subjectInput}
                onChange={(e) => setSubjectInput(e.target.value)}
                placeholder="例: 国語"
              />
              <button
                className="btn"
                type="button"
                onClick={() => {
                  const value = subjectInput.trim()
                  if (!value || data.subjects.includes(value)) {
                    return
                  }
                  void update((current) => ({ ...current, subjects: [...current.subjects, value] }))
                  setSubjectInput('')
                }}
              >
                追加
              </button>
            </div>
            <div className="row">
              {data.subjects.map((subject) => (
                <span key={subject} className="badge ok">
                  {subject}
                </span>
              ))}
            </div>
          </div>

          <div className="panel">
            <h3>先生登録</h3>
            <div className="row">
              <input
                value={teacherName}
                onChange={(e) => setTeacherName(e.target.value)}
                placeholder="先生名"
              />
              <input
                value={teacherSubjectsText}
                onChange={(e) => setTeacherSubjectsText(e.target.value)}
                placeholder="担当科目(カンマ区切り)"
              />
              <input
                value={teacherMemo}
                onChange={(e) => setTeacherMemo(e.target.value)}
                placeholder="メモ"
              />
              <button className="btn" type="button" onClick={() => void addTeacher()}>
                追加
              </button>
            </div>
            <table className="table">
              <thead>
                <tr>
                  <th>名前</th>
                  <th>科目</th>
                  <th>メモ</th>
                  <th>希望URL</th>
                </tr>
              </thead>
              <tbody>
                {data.teachers.map((teacher) => (
                  <tr key={teacher.id}>
                    <td>{teacher.name}</td>
                    <td>{teacher.subjects.join(', ')}</td>
                    <td>{teacher.memo}</td>
                    <td>
                      <Link to={`/availability/${sessionId}/teacher/${teacher.id}`}>入力ページ</Link>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <h3>生徒登録</h3>
            <div className="row">
              <input
                value={studentName}
                onChange={(e) => setStudentName(e.target.value)}
                placeholder="生徒名"
              />
              <input
                value={studentGrade}
                onChange={(e) => setStudentGrade(e.target.value)}
                placeholder="学年"
              />
              <select
                multiple
                value={studentSubjects}
                onChange={(e: ChangeEvent<HTMLSelectElement>) =>
                  setStudentSubjects(Array.from(e.target.selectedOptions, (option) => option.value))
                }
                style={{ minHeight: '60px' }}
              >
                {data.subjects.map((subject) => (
                  <option key={subject} value={subject}>
                    {subject}
                  </option>
                ))}
              </select>
              <input
                value={studentSubjectSlotsText}
                onChange={(e) => setStudentSubjectSlotsText(e.target.value)}
                placeholder="希望コマ数(例: 数学:5, 英語:3)"
              />
              <input
                value={studentUnavailableDatesText}
                onChange={(e) => setStudentUnavailableDatesText(e.target.value)}
                placeholder="出席不可能日(例: 2026-07-25, 2026-07-28)"
              />
              <input
                value={studentMemo}
                onChange={(e) => setStudentMemo(e.target.value)}
                placeholder="メモ"
              />
              <button className="btn" type="button" onClick={() => void addStudent()}>
                追加
              </button>
            </div>
            <table className="table">
              <thead>
                <tr>
                  <th>名前</th>
                  <th>学年</th>
                  <th>科目</th>
                  <th>希望コマ数</th>
                  <th>不可日数</th>
                  <th>メモ</th>
                  <th>希望URL</th>
                </tr>
              </thead>
              <tbody>
                {data.students.map((student) => (
                  <tr key={student.id}>
                    <td>{student.name}</td>
                    <td>{student.grade}</td>
                    <td>{student.subjects.join(', ')}</td>
                    <td>
                      {Object.entries(student.subjectSlots)
                        .map(([subject, count]) => `${subject}:${count}`)
                        .join(', ') || '-'}
                    </td>
                    <td>{student.unavailableDates.length}日</td>
                    <td>{student.memo}</td>
                    <td>
                      <Link to={`/availability/${sessionId}/student/${student.id}`}>入力ページ</Link>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <h3>先生×生徒 制約</h3>
            <div className="row">
              <select
                value={constraintTeacherId}
                onChange={(e) => setConstraintTeacherId(e.target.value)}
              >
                <option value="">先生を選択</option>
                {data.teachers.map((teacher) => (
                  <option key={teacher.id} value={teacher.id}>
                    {teacher.name}
                  </option>
                ))}
              </select>
              <select
                value={constraintStudentId}
                onChange={(e) => setConstraintStudentId(e.target.value)}
              >
                <option value="">生徒を選択</option>
                {data.students.map((student) => (
                  <option key={student.id} value={student.id}>
                    {student.name}
                  </option>
                ))}
              </select>
              <select value={constraintType} onChange={(e) => setConstraintType(e.target.value as ConstraintType)}>
                <option value="incompatible">組み合わせ不可</option>
                <option value="recommended">組み合わせ推奨</option>
              </select>
              <button className="btn" type="button" onClick={() => void upsertConstraint()}>
                保存
              </button>
            </div>
            <table className="table">
              <thead>
                <tr>
                  <th>先生</th>
                  <th>生徒</th>
                  <th>種別</th>
                </tr>
              </thead>
              <tbody>
                {data.constraints.map((constraint) => (
                  <tr key={constraint.id}>
                    <td>{data.teachers.find((teacher) => teacher.id === constraint.teacherId)?.name ?? '-'}</td>
                    <td>{data.students.find((student) => student.id === constraint.studentId)?.name ?? '-'}</td>
                    <td>
                      {constraint.type === 'incompatible' ? (
                        <span className="badge warn">不可</span>
                      ) : (
                        <span className="badge rec">推奨</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <h3>通常授業管理</h3>
            <div className="row">
              <select
                value={regularTeacherId}
                onChange={(e) => setRegularTeacherId(e.target.value)}
              >
                <option value="">先生を選択</option>
                {data.teachers.map((teacher) => (
                  <option key={teacher.id} value={teacher.id}>
                    {teacher.name}
                  </option>
                ))}
              </select>
              <select
                multiple
                value={regularStudentIds}
                onChange={(e) => setRegularStudentIds(Array.from(e.target.selectedOptions, (option) => option.value))}
                style={{ minHeight: '80px' }}
              >
                {data.students.map((student) => (
                  <option key={student.id} value={student.id}>
                    {student.name}
                  </option>
                ))}
              </select>
              <input
                value={regularSubject}
                onChange={(e) => setRegularSubject(e.target.value)}
                placeholder="科目"
              />
              <select
                value={regularDayOfWeek}
                onChange={(e) => setRegularDayOfWeek(e.target.value)}
              >
                <option value="">曜日を選択</option>
                <option value="0">日曜</option>
                <option value="1">月曜</option>
                <option value="2">火曜</option>
                <option value="3">水曜</option>
                <option value="4">木曜</option>
                <option value="5">金曜</option>
                <option value="6">土曜</option>
              </select>
              <input
                type="number"
                value={regularSlotNumber}
                onChange={(e) => setRegularSlotNumber(e.target.value)}
                placeholder="時限番号"
                min="1"
              />
              <button className="btn" type="button" onClick={() => void addRegularLesson()}>
                追加
              </button>
            </div>
            <p className="muted">通常授業は該当する曜日・時限のスロットに最優先で割り当てられます。</p>
            <table className="table">
              <thead>
                <tr>
                  <th>先生</th>
                  <th>生徒</th>
                  <th>科目</th>
                  <th>曜日</th>
                  <th>時限</th>
                  <th>操作</th>
                </tr>
              </thead>
              <tbody>
                {data.regularLessons.map((lesson) => {
                  const dayNames = ['日', '月', '火', '水', '木', '金', '土']
                  return (
                    <tr key={lesson.id}>
                      <td>{data.teachers.find((t) => t.id === lesson.teacherId)?.name ?? '-'}</td>
                      <td>
                        {lesson.studentIds
                          .map((id) => data.students.find((s) => s.id === id)?.name ?? '-')
                          .join(', ')}
                      </td>
                      <td>{lesson.subject}</td>
                      <td>{dayNames[lesson.dayOfWeek]}曜</td>
                      <td>{lesson.slotNumber}限</td>
                      <td>
                        <button
                          className="btn secondary"
                          type="button"
                          onClick={() => void removeRegularLesson(lesson.id)}
                        >
                          削除
                        </button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          <div className="panel">
            <div className="row">
              <h3>コマ割り</h3>
              <button className="btn secondary" type="button" onClick={() => void applyAutoAssign()}>
                自動提案(未割当)
              </button>
              <button className="btn" type="button" onClick={() => void applyTemplateAndAutoAssign()}>
                テストデータで自動提案
              </button>
            </div>
            <p className="muted">不可ペアは選択不可。推奨ペアを優先。先生1人 + 生徒1〜2人。同じコマに複数ペア可。</p>
            <div className="grid-slots">
              {slotKeys.map((slot) => {
                const slotAssignments = data.assignments[slot] ?? []
                const usedTeacherIds = new Set(slotAssignments.map((a) => a.teacherId).filter(Boolean))

                return (
                  <div className="slot-card" key={slot}>
                    <div className="slot-title">{slotLabel(slot)}</div>
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

                        return (
                          <div key={idx} className="assignment-block">
                            <select
                              value={assignment.teacherId}
                              onChange={(e) => void setSlotTeacher(slot, idx, e.target.value)}
                            >
                              <option value="">先生を選択</option>
                              {data.teachers.map((teacher) => {
                                const available = hasAvailability(data.availability, 'teacher', teacher.id, slot)
                                const usedElsewhere = usedTeacherIds.has(teacher.id) && teacher.id !== assignment.teacherId
                                return (
                                  <option key={teacher.id} value={teacher.id} disabled={!available || usedElsewhere}>
                                    {teacher.name} {!available ? '(希望なし)' : usedElsewhere ? '(割当済)' : ''}
                                  </option>
                                )
                              })}
                            </select>

                            {assignment.teacherId && (
                              <>
                                <select
                                  value={assignment.subject}
                                  onChange={(e) => void setSlotSubject(slot, idx, e.target.value)}
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
                                    <select
                                      key={pos}
                                      value={assignment.studentIds[pos] ?? ''}
                                      onChange={(e) => void setSlotStudent(slot, idx, pos, e.target.value)}
                                    >
                                      <option value="">{`生徒${pos + 1}を選択`}</option>
                                      {data.students.map((student) => {
                                        const available = hasAvailability(data.availability, 'student', student.id, slot)
                                        const tag = constraintFor(data.constraints, assignment.teacherId, student.id)
                                        const usedInOther = slotAssignments.some(
                                          (a, i) => i !== idx && a.studentIds.includes(student.id),
                                        )
                                        const isSelectedInOtherPosition = student.id === otherStudentId
                                        const disabled = !available || tag === 'incompatible' || usedInOther || isSelectedInOtherPosition
                                        const remaining = getTotalRemainingSlots(data.assignments, student)
                                        const tagLabel = tag === 'incompatible' ? ' [不可]' : tag === 'recommended' ? ' [推奨]' : ''
                                        const statusLabel = !available ? ' (希望なし)' : usedInOther ? ' (他ペア)' : ''

                                        return (
                                          <option key={student.id} value={student.id} disabled={disabled}>
                                            {student.name} 残{remaining}コマ{tagLabel}{statusLabel}
                                          </option>
                                        )
                                      })}
                                    </select>
                                  )
                                })}
                              </>
                            )}
                            {slotAssignments.length > 1 && (
                              <button
                                className="btn secondary"
                                type="button"
                                style={{ fontSize: '0.8em', marginTop: '4px' }}
                                onClick={() => void setSlotTeacher(slot, idx, '')}
                              >
                                このペアを削除
                              </button>
                            )}
                          </div>
                        )
                      })}
                      <button
                        className="btn secondary"
                        type="button"
                        style={{ fontSize: '0.8em', marginTop: '4px' }}
                        onClick={() => void addSlotAssignment(slot)}
                      >
                        ＋ ペア追加
                      </button>
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

// Helper: Check if date has regular lesson for student
const hasRegularLessonOnDate = (
  date: string,
  studentId: string,
  regularLessons: RegularLesson[],
): { hasLesson: boolean; lessonInfo?: string } => {
  const dateObj = new Date(`${date}T00:00:00`)
  const dayOfWeek = dateObj.getDay()

  const lesson = regularLessons.find(
    (lesson) => lesson.dayOfWeek === dayOfWeek && lesson.studentIds.includes(studentId),
  )

  if (lesson) {
    const dayNames = ['日', '月', '火', '水', '木', '金', '土']
    return {
      hasLesson: true,
      lessonInfo: `${lesson.subject} ${dayNames[dayOfWeek]}曜${lesson.slotNumber}限`,
    }
  }

  return { hasLesson: false }
}

// Teacher Input Component
const TeacherInputPage = ({
  sessionId,
  data,
  teacher,
}: {
  sessionId: string
  data: SessionData
  teacher: Teacher
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const [localAvailability, setLocalAvailability] = useState<Set<string>>(() => {
    const key = personKey('teacher', teacher.id)
    return new Set(data.availability[key] ?? [])
  })
  const [submitting, setSubmitting] = useState(false)

  const toggleSlot = (date: string, slotNum: number) => {
    const slotKey = `${date}_${slotNum}`
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

  const handleSubmit = () => {
    setSubmitting(true)
    const key = personKey('teacher', teacher.id)
    const next: SessionData = {
      ...data,
      availability: {
        ...data.availability,
        [key]: Array.from(localAvailability),
      },
    }
    saveSession(sessionId, next).catch(() => {})
    navigate(`/complete/${sessionId}`)
  }

  return (
    <div className="availability-container">
      <div className="availability-header">
        <h2>{data.settings.name} - 先生希望入力</h2>
        <p>
          対象: <strong>{teacher.name}</strong>
        </p>
        <p className="muted">出席可能なコマをタップして選択してください。</p>
      </div>

      <div className="teacher-table-wrapper">
        <table className="teacher-table">
          <thead>
            <tr>
              <th className="date-header">日付</th>
              {Array.from({ length: data.settings.slotsPerDay }, (_, i) => (
                <th key={i}>{i + 1}限</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {dates.map((date) => (
              <tr key={date}>
                <td className="date-cell">{date}</td>
                {Array.from({ length: data.settings.slotsPerDay }, (_, i) => {
                  const slotNum = i + 1
                  const slotKey = `${date}_${slotNum}`
                  const isOn = localAvailability.has(slotKey)
                  return (
                    <td key={slotNum}>
                      <button
                        className={`teacher-slot-btn ${isOn ? 'active' : ''}`}
                        onClick={() => toggleSlot(date, slotNum)}
                        type="button"
                      >
                        {isOn ? '○' : ''}
                      </button>
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="submit-section">
        <button
          className="submit-btn"
          onClick={handleSubmit}
          disabled={submitting}
          type="button"
        >
          {submitting ? '送信中...' : '送信'}
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
}: {
  sessionId: string
  data: SessionData
  student: Student
}) => {
  const navigate = useNavigate()
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const [selectedSubjects, setSelectedSubjects] = useState<Set<string>>(
    new Set(student.subjects ?? []),
  )
  const [subjectSlots, setSubjectSlots] = useState<Record<string, number>>(
    student.subjectSlots ?? {},
  )
  const [unavailableDates, setUnavailableDates] = useState<Set<string>>(
    new Set(student.unavailableDates ?? []),
  )
  const [submitting, setSubmitting] = useState(false)

  const toggleSubject = (subject: string) => {
    setSelectedSubjects((prev) => {
      const next = new Set(prev)
      if (next.has(subject)) {
        next.delete(subject)
      } else {
        next.add(subject)
      }
      return next
    })
  }

  const toggleDate = (date: string) => {
    const regularCheck = hasRegularLessonOnDate(date, student.id, data.regularLessons)

    if (regularCheck.hasLesson && !unavailableDates.has(date)) {
      const confirmed = window.confirm(
        `この日は通常授業（${regularCheck.lessonInfo}）がありますが、出席不可としますか？`,
      )
      if (!confirmed) {
        return
      }
    }

    setUnavailableDates((prev) => {
      const next = new Set(prev)
      if (next.has(date)) {
        next.delete(date)
      } else {
        next.add(date)
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
    setSubmitting(true)
    const subjects = Array.from(selectedSubjects)
    const updatedStudents = data.students.map((s) =>
      s.id === student.id
        ? {
            ...s,
            subjects,
            subjectSlots,
            unavailableDates: Array.from(unavailableDates),
            submittedAt: Date.now(),
          }
        : s,
    )

    const next: SessionData = {
      ...data,
      students: updatedStudents,
    }
    saveSession(sessionId, next).catch(() => {})
    navigate(`/complete/${sessionId}`)
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
        <h3>希望科目</h3>
        <p className="muted">受講を希望する科目を選択してください。</p>
        <div className="subject-checkboxes">
          {data.subjects.map((subject) => (
            <label className="row" key={subject}>
              <input
                type="checkbox"
                checked={selectedSubjects.has(subject)}
                onChange={() => toggleSubject(subject)}
              />
              <span>{subject}</span>
            </label>
          ))}
        </div>
      </div>

      <div className="student-form-section">
        <h3>希望科目のコマ数</h3>
        <p className="muted">各科目について、希望するコマ数を入力してください。</p>
        <div className="subject-slots-form">
          {data.subjects.filter((s) => selectedSubjects.has(s)).map((subject) => (
            <div key={subject} className="form-row">
              <label htmlFor={`subject-${subject}`}>{subject}:</label>
              <input
                id={`subject-${subject}`}
                type="number"
                min="0"
                value={subjectSlots[subject] ?? 0}
                onChange={(e) => handleSubjectSlotsChange(subject, e.target.value)}
              />
              <span className="form-unit">コマ</span>
            </div>
          ))}
        </div>
      </div>

      <div className="student-form-section">
        <h3>出席不可日</h3>
        <p className="muted">出席できない日をタップして選択してください。</p>
        <div className="date-checkboxes">
          {dates.map((date) => {
            const isUnavailable = unavailableDates.has(date)
            const regularCheck = hasRegularLessonOnDate(date, student.id, data.regularLessons)
            return (
              <div key={date} className="date-checkbox-item">
                <button
                  className={`date-checkbox-btn ${isUnavailable ? 'checked' : ''}`}
                  onClick={() => toggleDate(date)}
                  type="button"
                >
                  <span className="checkbox-icon">{isUnavailable ? '✓' : ''}</span>
                  <span className="date-label">{date}</span>
                  {regularCheck.hasLesson && (
                    <span className="regular-lesson-badge">通常授業</span>
                  )}
                </button>
              </div>
            )
          })}
        </div>
      </div>

      <div className="submit-section">
        <button
          className="submit-btn"
          onClick={handleSubmit}
          disabled={submitting}
          type="button"
        >
          {submitting ? '送信中...' : '送信'}
        </button>
      </div>
    </div>
  )
}

const AvailabilityPage = () => {
  const { sessionId = 'main', personType = 'teacher', personId = '' } = useParams()
  const [data, setData] = useState<SessionData | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    setLoading(true)
    const unsub = watchSession(sessionId, (value) => {
      setData(value)
      setLoading(false)
    })
    return () => unsub()
  }, [sessionId])

  const currentPerson = useMemo(() => {
    if (!data) {
      return null
    }
    if (personType === 'teacher') {
      return data.teachers.find((teacher) => teacher.id === personId) ?? null
    }
    return data.students.find((student) => student.id === personId) ?? null
  }, [data, personId, personType])

  if (loading) {
    return (
      <div className="app-shell">
        <div className="panel">読み込み中...</div>
      </div>
    )
  }

  if (!data || !currentPerson) {
    return (
      <div className="app-shell">
        <div className="panel">
          入力対象が見つかりません。管理者にURLを確認してください。
          <br />
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  if (personType === 'teacher') {
    // Runtime type check for Teacher
    if ('subjects' in currentPerson && Array.isArray(currentPerson.subjects)) {
      return <TeacherInputPage sessionId={sessionId} data={data} teacher={currentPerson as Teacher} />
    }
  } else if (personType === 'student') {
    // Runtime type check for Student
    if ('grade' in currentPerson && 'subjectSlots' in currentPerson) {
      return <StudentInputPage sessionId={sessionId} data={data} student={currentPerson as Student} />
    }
  }

  // Fallback if person type doesn't match
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
  const { sessionId = 'main' } = useParams()

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>入力完了</h2>
        <p>データの送信が完了しました。ありがとうございます。</p>
        <div className="row">
          <Link className="btn" to={`/admin/${sessionId}`} state={{ skipAuth: true }}>設定に戻る</Link>
        </div>
      </div>
    </div>
  )
}

function App() {
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
