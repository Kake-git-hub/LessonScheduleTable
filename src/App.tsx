import { useEffect, useMemo, useState } from 'react'
import { Link, Route, Routes, useNavigate, useParams } from 'react-router-dom'
import './App.css'
import { loadSession, saveSession, watchSession } from './firebase'
import type {
  Assignment,
  ConstraintType,
  PairConstraint,
  PersonType,
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
    { id: 's001', name: '青木 太郎', grade: '中3', subjects: ['数学', '英語'], memo: '受験対策' },
    { id: 's002', name: '伊藤 花', grade: '中2', subjects: ['英語', '国語'], memo: '' },
    { id: 's003', name: '上田 陽介', grade: '高1', subjects: ['数学', '理科'], memo: '' },
    { id: 's004', name: '岡本 美咲', grade: '高2', subjects: ['英語', '数学'], memo: '' },
    { id: 's005', name: '加藤 駿', grade: '中3', subjects: ['国語', '英語'], memo: '' },
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

  return {
    settings,
    subjects,
    teachers,
    students,
    constraints,
    availability,
    assignments: {},
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

    let bestPlan: { score: number; assignment: Assignment } | null = null

    const teachers = data.teachers.filter((teacher) =>
      hasAvailability(data.availability, 'teacher', teacher.id, slot),
    )

    for (const teacher of teachers) {
      const candidates = data.students.filter((student) => {
        if (!hasAvailability(data.availability, 'student', student.id, slot)) {
          return false
        }
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

        const recommendScore = combo.reduce((score, student) => {
          return score + (constraintFor(data.constraints, teacher.id, student.id) === 'recommended' ? 30 : 0)
        }, 0)

        const studentLoadPenalty = combo.reduce((score, student) => {
          return score + countStudentLoad(nextAssignments, student.id) * 8
        }, 0)

        const groupBonus = combo.length === 2 ? 5 : 0
        const score = 100 + recommendScore + groupBonus - teacherLoad * 6 - studentLoadPenalty

        if (!bestPlan || score > bestPlan.score) {
          bestPlan = {
            score,
            assignment: {
              teacherId: teacher.id,
              studentIds: combo.map((student) => student.id),
              subject: commonSubjects[0],
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
  const [studentMemo, setStudentMemo] = useState('')

  const [constraintTeacherId, setConstraintTeacherId] = useState('')
  const [constraintStudentId, setConstraintStudentId] = useState('')
  const [constraintType, setConstraintType] = useState<ConstraintType>('incompatible')

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
    const student: Student = {
      id: createId(),
      name: studentName.trim(),
      grade: studentGrade.trim(),
      subjects: parseList(studentSubjectsText),
      memo: studentMemo.trim(),
    }

    await update((current) => ({ ...current, students: [...current.students, student] }))
    setStudentName('')
    setStudentGrade('')
    setStudentSubjectsText('')
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

  const slotKeys = useMemo(() => (data ? buildSlotKeys(data.settings) : []), [data])

  const currentPerson = useMemo(() => {
    if (!data) {
      return null
    }
    if (personType === 'teacher') {
      return data.teachers.find((teacher) => teacher.id === personId) ?? null
    }
    return data.students.find((student) => student.id === personId) ?? null
  }, [data, personId, personType])

  const selected =
    data?.availability[personKey(personType as PersonType, personId)] ?? []

  const toggle = async (slot: string): Promise<void> => {
    if (!data || !personId) {
      return
    }
    const key = personKey(personType as PersonType, personId)
    const set = new Set(data.availability[key] ?? [])
    if (set.has(slot)) {
      set.delete(slot)
    } else {
      set.add(slot)
    }

    const next: SessionData = {
      ...data,
      availability: {
        ...data.availability,
        [key]: Array.from(set),
      },
    }

    setData(next)
    await saveSession(sessionId, next)
  }

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
          入力対象が見つかりません。管理者にURLを確認してください。<br />
          <Link to="/">ホームに戻る</Link>
        </div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <div className="panel">
        <h2>
          {data.settings.name} - {personType === 'teacher' ? '先生' : '生徒'}希望入力
        </h2>
        <p>
          対象: <strong>{currentPerson.name}</strong>
        </p>
        <p className="muted">参加できるコマをONにしてください。</p>
      </div>

      <div className="panel">
        <div className="availability-grid">
          {slotKeys.map((slot) => {
            const on = selected.includes(slot)
            return (
              <button
                className={`slot-toggle ${on ? 'on' : ''}`}
                key={slot}
                onClick={() => void toggle(slot)}
                type="button"
              >
                {slotLabel(slot)}
              </button>
            )
          })}
        </div>
      </div>
    </div>
  )
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
