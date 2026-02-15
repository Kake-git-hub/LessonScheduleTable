import { useEffect, useMemo, useState } from 'react'
import { Link, Route, Routes, useNavigate, useParams } from 'react-router-dom'
import './App.css'
import { loadSession, saveSession, watchSession } from './firebase'
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
    endDate: '2026-07-31',
    slotsPerDay: 5,
    holidays: ['2026-07-26'],
  }

  const subjects = ['数学', '英語', '国語', '理科']

  const teachers: Teacher[] = [
    { id: 't001', name: '田中先生', subjects: ['数学', '理科'], memo: '中3理系担当' },
    { id: 't002', name: '佐藤先生', subjects: ['英語', '国語'], memo: '文系担当' },
    { id: 't003', name: '鈴木先生', subjects: ['数学', '英語'], memo: '高1〜高2担当' },
  ]

  const students: Student[] = [
    {
      id: 's001',
      name: '青木 太郎',
      grade: '中3',
      subjects: ['数学', '英語'],
      subjectSlots: { 数学: 5, 英語: 3 },
      unavailableDates: ['2026-07-25'],
      memo: '受験対策',
      submittedAt: Date.now() - 4000,
    },
    {
      id: 's002',
      name: '伊藤 花',
      grade: '中2',
      subjects: ['英語', '国語'],
      subjectSlots: { 英語: 4, 国語: 2 },
      unavailableDates: [],
      memo: '',
      submittedAt: Date.now() - 3000,
    },
    {
      id: 's003',
      name: '上田 陽介',
      grade: '高1',
      subjects: ['数学', '理科'],
      subjectSlots: { 数学: 3, 理科: 3 },
      unavailableDates: ['2026-07-22', '2026-07-29'],
      memo: '',
      submittedAt: Date.now() - 2000,
    },
    {
      id: 's004',
      name: '岡本 美咲',
      grade: '高2',
      subjects: ['英語', '数学'],
      subjectSlots: { 英語: 4, 数学: 4 },
      unavailableDates: [],
      memo: '',
      submittedAt: Date.now() - 1000,
    },
    {
      id: 's005',
      name: '加藤 駿',
      grade: '中3',
      subjects: ['国語', '英語'],
      subjectSlots: { 国語: 3, 英語: 2 },
      unavailableDates: ['2026-07-28'],
      memo: '',
      submittedAt: Date.now(),
    },
  ]

  const constraints: PairConstraint[] = [
    { id: 'c001', teacherId: 't001', studentId: 's002', type: 'incompatible' },
    { id: 'c002', teacherId: 't002', studentId: 's003', type: 'incompatible' },
    { id: 'c003', teacherId: 't001', studentId: 's003', type: 'recommended' },
    { id: 'c004', teacherId: 't002', studentId: 's005', type: 'recommended' },
    { id: 'c005', teacherId: 't003', studentId: 's004', type: 'recommended' },
  ]

  const slotKeys = buildSlotKeys(settings)
  const availability: SessionData['availability'] = {
    [personKey('teacher', 't001')]: slotKeys.filter((slot) => /_(1|2|3)$/.test(slot)),
    [personKey('teacher', 't002')]: slotKeys.filter((slot) => /_(2|3|4)$/.test(slot)),
    [personKey('teacher', 't003')]: slotKeys.filter((slot) => /_(3|4|5)$/.test(slot)),
    [personKey('student', 's001')]: slotKeys.filter((slot) => /_(1|2|4)$/.test(slot)),
    [personKey('student', 's002')]: slotKeys.filter((slot) => /_(2|3|5)$/.test(slot)),
    [personKey('student', 's003')]: slotKeys.filter((slot) => /_(1|3|4)$/.test(slot)),
    [personKey('student', 's004')]: slotKeys.filter((slot) => /_(2|4|5)$/.test(slot)),
    [personKey('student', 's005')]: slotKeys.filter((slot) => /_(1|2|3)$/.test(slot)),
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
    {
      id: 'r002',
      teacherId: 't002',
      studentIds: ['s002', 's005'],
      subject: '国語',
      dayOfWeek: 3,
      slotNumber: 2,
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

const countTeacherLoad = (assignments: Record<string, Assignment>, teacherId: string): number =>
  Object.values(assignments).filter((assignment) => assignment.teacherId === teacherId).length

const countStudentLoad = (assignments: Record<string, Assignment>, studentId: string): number =>
  Object.values(assignments).filter((assignment) => assignment.studentIds.includes(studentId)).length

const countStudentSubjectLoad = (
  assignments: Record<string, Assignment>,
  studentId: string,
  subject: string,
): number =>
  Object.values(assignments).filter(
    (assignment) => assignment.studentIds.includes(studentId) && assignment.subject === subject,
  ).length

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
): Record<string, Assignment> => {
  const nextAssignments: Record<string, Assignment> = onlyEmpty ? { ...data.assignments } : {}

  for (const slot of slots) {
    if (onlyEmpty && nextAssignments[slot]) {
      continue
    }

    // Rule 2: Check if there's a regular lesson for this slot
    const regularLesson = findRegularLessonForSlot(data.regularLessons, slot)
    if (regularLesson) {
      nextAssignments[slot] = {
        teacherId: regularLesson.teacherId,
        studentIds: regularLesson.studentIds,
        subject: regularLesson.subject,
      }
      continue
    }

    let bestPlan: { score: number; assignment: Assignment } | null = null

    const teachers = data.teachers.filter((teacher) =>
      hasAvailability(data.availability, 'teacher', teacher.id, slot),
    )

    for (const teacher of teachers) {
      const candidates = data.students.filter((student) => {
        // Check teacher availability (existing)
        if (!hasAvailability(data.availability, 'student', student.id, slot)) {
          return false
        }
        // Rule 1: Check unavailable dates
        if (!isStudentAvailable(student, slot)) {
          return false
        }
        // Check incompatibility (existing)
        if (constraintFor(data.constraints, teacher.id, student.id) === 'incompatible') {
          return false
        }
        return teacher.subjects.some((subject) => student.subjects.includes(subject))
      })

      if (candidates.length === 0) {
        continue
      }

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

        if (commonSubjects.length === 0) {
          continue
        }

        // Rule 3: Filter subjects based on subject slot requests
        const viableSubjects = commonSubjects.filter((subject) =>
          combo.every((student) => {
            const requested = student.subjectSlots[subject] ?? 0
            const allocated = countStudentSubjectLoad(nextAssignments, student.id, subject)
            return allocated < requested
          }),
        )

        if (viableSubjects.length === 0) {
          continue
        }

        const recommendScore = combo.reduce((score, student) => {
          return score + (constraintFor(data.constraints, teacher.id, student.id) === 'recommended' ? 30 : 0)
        }, 0)

        const studentLoadPenalty = combo.reduce((score, student) => {
          return score + countStudentLoad(nextAssignments, student.id) * 8
        }, 0)

        // Rule 3: Priority bonus based on student index (submission order)
        const priorityBonus = combo.reduce((bonus, student) => {
          const studentIndex = data.students.findIndex((s) => s.id === student.id)
          return bonus + Math.max(0, 20 - studentIndex * 2)
        }, 0)

        // Rule 3: Bonus for unfulfilled subject requests
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
    }

    if (bestPlan) {
      nextAssignments[slot] = bestPlan.assignment
    }
  }

  return nextAssignments
}

const HomePage = () => {
  const navigate = useNavigate()
  const [sessionId, setSessionId] = useState('main')

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>講習コマ割りアプリ</h2>
        <p className="muted">管理画面と希望入力URLを分けて運用できます。</p>
        <div className="row">
          <input value={sessionId} onChange={(e) => setSessionId(e.target.value.trim())} />
          <button
            className="btn"
            onClick={() => navigate(`/admin/${sessionId || 'main'}`)}
            type="button"
          >
            管理画面を開く
          </button>
        </div>
      </div>
    </div>
  )
}

const AdminPage = () => {
  const { sessionId = 'main' } = useParams()
  const { data, setData, loading } = useSessionData(sessionId)
  const [passwordInput, setPasswordInput] = useState('')
  const [authorized, setAuthorized] = useState(false)

  const [subjectInput, setSubjectInput] = useState('')
  const [teacherName, setTeacherName] = useState('')
  const [teacherSubjectsText, setTeacherSubjectsText] = useState('')
  const [teacherMemo, setTeacherMemo] = useState('')

  const [studentName, setStudentName] = useState('')
  const [studentGrade, setStudentGrade] = useState('')
  const [studentSubjectsText, setStudentSubjectsText] = useState('')
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
    setAuthorized(false)
    setPasswordInput('')
  }, [sessionId])

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

  const login = (): void => {
    if (!data) {
      return
    }
    setAuthorized(passwordInput === data.settings.adminPassword)
  }

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
      subjects: parseList(studentSubjectsText),
      subjectSlots,
      unavailableDates,
      memo: studentMemo.trim(),
      submittedAt: Date.now(),
    }

    await update((current) => ({ ...current, students: [...current.students, student] }))
    setStudentName('')
    setStudentGrade('')
    setStudentSubjectsText('')
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

  const setSlotTeacher = async (slot: string, teacherId: string): Promise<void> => {
    await update((current) => {
      const prev = current.assignments[slot]
      if (!teacherId) {
        const nextAssignments = { ...current.assignments }
        delete nextAssignments[slot]
        return { ...current, assignments: nextAssignments }
      }

      const currentTeacher = current.teachers.find((item) => item.id === teacherId)
      const nextSubject =
        prev?.subject && currentTeacher?.subjects.includes(prev.subject)
          ? prev.subject
          : (currentTeacher?.subjects[0] ?? '')

      return {
        ...current,
        assignments: {
          ...current.assignments,
          [slot]: {
            teacherId,
            studentIds: prev?.teacherId === teacherId ? prev.studentIds : [],
            subject: nextSubject,
          },
        },
      }
    })
  }

  const toggleSlotStudent = async (slot: string, studentId: string): Promise<void> => {
    await update((current) => {
      const assignment = current.assignments[slot]
      if (!assignment) {
        return current
      }

      const already = assignment.studentIds.includes(studentId)
      let studentIds = already
        ? assignment.studentIds.filter((id) => id !== studentId)
        : [...assignment.studentIds, studentId]

      if (studentIds.length > 2) {
        studentIds = studentIds.slice(0, 2)
      }

      const teacher = current.teachers.find((item) => item.id === assignment.teacherId)
      const commonSubjects = (teacher?.subjects ?? []).filter((subject) =>
        studentIds.every((id) => current.students.find((student) => student.id === id)?.subjects.includes(subject)),
      )

      return {
        ...current,
        assignments: {
          ...current.assignments,
          [slot]: {
            ...assignment,
            studentIds,
            subject: commonSubjects.includes(assignment.subject)
              ? assignment.subject
              : (commonSubjects[0] ?? assignment.subject),
          },
        },
      }
    })
  }

  const setSlotSubject = async (slot: string, subject: string): Promise<void> => {
    await update((current) => {
      const assignment = current.assignments[slot]
      if (!assignment) {
        return current
      }
      return {
        ...current,
        assignments: {
          ...current.assignments,
          [slot]: { ...assignment, subject },
        },
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
          <h3>管理者パスワード</h3>
          <div className="row">
            <input
              type="password"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              placeholder="パスワード"
            />
            <button className="btn" type="button" onClick={login}>
              ログイン
            </button>
          </div>
          <p className="muted">初期値は admin1234 です。ログイン後に必ず変更してください。</p>
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
              <input
                value={studentSubjectsText}
                onChange={(e) => setStudentSubjectsText(e.target.value)}
                placeholder="受講科目(カンマ区切り)"
              />
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
            <p className="muted">不可ペアは選択不可。推奨ペアを優先。先生1人 + 生徒1〜2人。</p>
            <div className="grid-slots">
              {slotKeys.map((slot) => {
                const assignment = data.assignments[slot]
                const selectedTeacher = data.teachers.find((teacher) => teacher.id === assignment?.teacherId)
                const selectedStudents = data.students.filter((student) =>
                  assignment?.studentIds.includes(student.id),
                )
                const subjectOptions = selectedTeacher
                  ? selectedTeacher.subjects.filter((subject) =>
                      selectedStudents.length === 0
                        ? true
                        : selectedStudents.every((student) => student.subjects.includes(subject)),
                    )
                  : []

                return (
                  <div className="slot-card" key={slot}>
                    <div className="slot-title">{slotLabel(slot)}</div>
                    <div className="list">
                      <select
                        value={assignment?.teacherId ?? ''}
                        onChange={(e) => void setSlotTeacher(slot, e.target.value)}
                      >
                        <option value="">先生を選択</option>
                        {data.teachers.map((teacher) => {
                          const available = hasAvailability(data.availability, 'teacher', teacher.id, slot)
                          return (
                            <option key={teacher.id} value={teacher.id} disabled={!available}>
                              {teacher.name} {available ? '' : '(希望なし)'}
                            </option>
                          )
                        })}
                      </select>

                      {assignment && (
                        <>
                          <select
                            value={assignment.subject}
                            onChange={(e) => void setSlotSubject(slot, e.target.value)}
                          >
                            {subjectOptions.map((subject) => (
                              <option key={subject} value={subject}>
                                {subject}
                              </option>
                            ))}
                          </select>

                          {data.students.map((student) => {
                            const available = hasAvailability(data.availability, 'student', student.id, slot)
                            const tag = constraintFor(data.constraints, assignment.teacherId, student.id)
                            const disabled =
                              !available ||
                              tag === 'incompatible' ||
                              (!assignment.studentIds.includes(student.id) && assignment.studentIds.length >= 2)
                            const checked = assignment.studentIds.includes(student.id)

                            return (
                              <label className="row" key={student.id}>
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  disabled={disabled}
                                  onChange={() => void toggleSlotStudent(slot, student.id)}
                                />
                                <span>{student.name}</span>
                                {tag === 'incompatible' ? (
                                  <span className="badge warn">不可</span>
                                ) : tag === 'recommended' ? (
                                  <span className="badge rec">推奨</span>
                                ) : null}
                                {!available ? <span className="muted">希望なし</span> : null}
                              </label>
                            )
                          })}
                        </>
                      )}
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

  for (let cursor = new Date(start); cursor <= end; cursor.setDate(cursor.getDate() + 1)) {
    const y = cursor.getFullYear()
    const m = String(cursor.getMonth() + 1).padStart(2, '0')
    const d = String(cursor.getDate()).padStart(2, '0')
    const iso = `${y}-${m}-${d}`
    if (!holidaySet.has(iso)) {
      dates.push(iso)
    }
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
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const [localAvailability, setLocalAvailability] = useState<Set<string>>(() => {
    const key = personKey('teacher', teacher.id)
    return new Set(data.availability[key] ?? [])
  })
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)

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

  const handleSubmit = async () => {
    setSubmitting(true)
    try {
      const key = personKey('teacher', teacher.id)
      const next: SessionData = {
        ...data,
        availability: {
          ...data.availability,
          [key]: Array.from(localAvailability),
        },
      }
      await saveSession(sessionId, next)
      setSubmitted(true)
      setTimeout(() => setSubmitted(false), 3000)
    } finally {
      setSubmitting(false)
    }
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
        {submitted && <p className="success-message">送信しました！</p>}
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
  const dates = useMemo(() => getDatesInRange(data.settings), [data.settings])
  const [subjectSlots, setSubjectSlots] = useState<Record<string, number>>(
    student.subjectSlots ?? {},
  )
  const [unavailableDates, setUnavailableDates] = useState<Set<string>>(
    new Set(student.unavailableDates ?? []),
  )
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)

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
    const numValue = Number.parseInt(value, 10)
    setSubjectSlots((prev) => ({
      ...prev,
      [subject]: Number.isNaN(numValue) ? 0 : numValue,
    }))
  }

  const handleSubmit = async () => {
    setSubmitting(true)
    try {
      const updatedStudents = data.students.map((s) =>
        s.id === student.id
          ? {
              ...s,
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
      await saveSession(sessionId, next)
      setSubmitted(true)
      setTimeout(() => setSubmitted(false), 3000)
    } finally {
      setSubmitting(false)
    }
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
        <h3>希望科目のコマ数</h3>
        <p className="muted">各科目について、希望するコマ数を入力してください。</p>
        <div className="subject-slots-form">
          {student.subjects.map((subject) => (
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
        {submitted && <p className="success-message">送信しました！</p>}
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
    return <TeacherInputPage sessionId={sessionId} data={data} teacher={currentPerson as Teacher} />
  }

  return <StudentInputPage sessionId={sessionId} data={data} student={currentPerson as Student} />
}

const BootPage = () => {
  const navigate = useNavigate()

  useEffect(() => {
    void (async () => {
      const existing = await loadSession('main')
      if (!existing) {
        await saveSession('main', emptySession())
      }
      navigate('/admin/main', { replace: true })
    })()
  }, [navigate])

  return (
    <div className="app-shell">
      <div className="panel">初期化中...</div>
    </div>
  )
}

function App() {
  return (
    <Routes>
      <Route path="/" element={<HomePage />} />
      <Route path="/boot" element={<BootPage />} />
      <Route path="/admin/:sessionId" element={<AdminPage />} />
      <Route path="/availability/:sessionId/:personType/:personId" element={<AvailabilityPage />} />
    </Routes>
  )
}

export default App
