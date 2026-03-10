import { buildIncrementalAutoAssignments, buildMendanAutoAssignments, dedupeMakeupInfosByOccurrence } from '../autoAssign'
import { blockReasonLabel } from '../autoAssignDiagnostics'
import type { SessionData, Student, Teacher } from '../../types'

const makeTeacher = (overrides: Partial<Teacher> = {}): Teacher => ({
  id: 't1',
  name: '田中先生',
  email: '',
  subjects: ['数', '英'],
  memo: '',
  ...overrides,
})

const makeStudent = (overrides: Partial<Student> = {}): Student => ({
  id: 's1',
  name: '山田太郎',
  email: '',
  grade: '中1',
  subjects: ['数'],
  subjectSlots: { '数': 2 },
  unavailableDates: [],
  preferredSlots: [],
  unavailableSlots: [],
  memo: '',
  submittedAt: 1000,
  ...overrides,
})

const makeSessionData = (overrides: Partial<SessionData> = {}): SessionData => ({
  settings: {
    name: '夏期講習',
    adminPassword: 'admin',
    startDate: '2026-07-21',
    endDate: '2026-07-22',
    slotsPerDay: 2,
    holidays: [],
  },
  subjects: ['数', '英'],
  managers: [],
  teachers: [makeTeacher()],
  students: [makeStudent()],
  constraints: [],
  gradeConstraints: [],
  availability: {
    'teacher:t1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
    'student:s1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
  },
  assignments: {},
  regularLessons: [],
  ...overrides,
})

describe('buildIncrementalAutoAssignments', () => {
  const slots = ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2']

  it('creates assignments for unassigned students', async () => {
    const data = makeSessionData()
    const result = await buildIncrementalAutoAssignments(data, slots)
    // Student requested 2 math slots, should get assigned
    const allAssigns = Object.values(result.assignments).flat()
    const studentAssigns = allAssigns.filter(a => a.studentIds.includes('s1'))
    expect(studentAssigns.length).toBe(2)
    expect(studentAssigns.every(a => a.subject === '数')).toBe(true)
  })

  it('respects teacher availability', async () => {
    const data = makeSessionData({
      availability: {
        'teacher:t1': ['2026-07-21_1'], // only available for 1 slot
        'student:s1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
      },
    })
    const result = await buildIncrementalAutoAssignments(data, slots)
    const allAssigns = Object.values(result.assignments).flat()
    const teacherAssigns = allAssigns.filter(a => a.teacherId === 't1')
    // Teacher should only be assigned on slots they're available
    expect(teacherAssigns.length).toBeLessThanOrEqual(1)
  })

  it('respects student availability', async () => {
    const data = makeSessionData({
      students: [makeStudent({ unavailableSlots: ['2026-07-21_1', '2026-07-21_2'] })],
    })
    const result = await buildIncrementalAutoAssignments(data, slots)
    // Student should not be assigned on unavailable slots
    expect(result.assignments['2026-07-21_1']?.some(a => a.studentIds.includes('s1'))).toBeFalsy()
    expect(result.assignments['2026-07-21_2']?.some(a => a.studentIds.includes('s1'))).toBeFalsy()
  })

  it('respects incompatible constraints', async () => {
    const data = makeSessionData({
      teachers: [makeTeacher()],
      students: [makeStudent(), makeStudent({ id: 's2', name: '佐藤花子', subjects: ['数'], subjectSlots: { '数': 2 }, submittedAt: 2000 })],
      constraints: [{ id: 'c1', personAId: 't1', personBId: 's2', personAType: 'teacher' as const, personBType: 'student' as const, type: 'incompatible' as const }],
      availability: {
        'teacher:t1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
        'student:s1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
        'student:s2': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
      },
    })
    const result = await buildIncrementalAutoAssignments(data, slots)
    const allAssigns = Object.values(result.assignments).flat()
    // s2 should never be assigned with t1 due to incompatible constraint
    const s2WithT1 = allAssigns.filter(a => a.teacherId === 't1' && a.studentIds.includes('s2'))
    expect(s2WithT1).toHaveLength(0)
  })

  it('cleans up assignments for deleted teachers', async () => {
    const data = makeSessionData({
      teachers: [makeTeacher({ id: 't2', name: '佐藤先生' })], // t1 deleted, t2 exists
      assignments: {
        '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }], // t1 no longer exists
      },
      availability: {
        'teacher:t2': ['2026-07-21_1', '2026-07-21_2'],
        'student:s1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
      },
    })
    const result = await buildIncrementalAutoAssignments(data, slots)
    // t1 assignment should be cleaned up (replaced or removed)
    const t1Assigns = Object.values(result.assignments).flat().filter(a => a.teacherId === 't1')
    expect(t1Assigns).toHaveLength(0)
  })

  it('removes students when they become unavailable', async () => {
    const data = makeSessionData({
      assignments: {
        '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      },
      students: [makeStudent({ unavailableSlots: ['2026-07-21_1'] })], // s1 no longer available for this slot
    })
    const result = await buildIncrementalAutoAssignments(data, slots)
    const slot1Assigns = result.assignments['2026-07-21_1'] ?? []
    const s1InSlot1 = slot1Assigns.some(a => a.studentIds.includes('s1'))
    expect(s1InSlot1).toBe(false)
  })

  it('handles empty session (no students)', async () => {
    const data = makeSessionData({ students: [] })
    const result = await buildIncrementalAutoAssignments(data, slots)
    const allAssigns = Object.values(result.assignments).flat()
    expect(allAssigns).toHaveLength(0)
  })

  it('reports diagnostics for students with unmet demand', async () => {
    const data = makeSessionData({
      teachers: [makeTeacher({ subjects: ['英'] })],
      students: [makeStudent({ subjects: ['数'], subjectSlots: { 数: 1 } })],
      availability: {
        'teacher:t1': ['2026-07-21_1'],
        'student:s1': ['2026-07-21_1'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1'])

    expect(result.diagnostics).toHaveLength(1)
    expect(result.diagnostics[0]).toMatchObject({
      studentId: 's1',
      subjectDemand: {
        数: { requested: 1, assigned: 0, remaining: 1 },
      },
      blockReasons: expect.objectContaining({
        subjectMismatch: 1,
      }),
      primaryBottleneck: 'subjectMismatch',
    })
    expect(blockReasonLabel(result.diagnostics[0].primaryBottleneck)).toBe('科目が合う講師がいない')
  })

  it('keeps regular substitute assignments on rerun even when special demand is zero', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 3, holidays: [] },
      teachers: [
        makeTeacher({ id: 't1', name: '鈴木先生', subjects: ['数', '英'] }),
        makeTeacher({ id: 't2', name: '山本先生', subjects: ['数', '英'] }),
      ],
      students: [
        makeStudent({ id: 's1', name: '上田陽介', subjects: [], subjectSlots: {}, submittedAt: 1000 }),
        makeStudent({ id: 's2', name: '同席生徒', subjects: ['英'], subjectSlots: { 英: 1 }, submittedAt: 2000 }),
      ],
      regularLessons: [
        { id: 'r1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 },
      ],
      assignments: {
        '2026-07-21_2': [{
          teacherId: 't2',
          studentIds: ['s2', 's1'],
          subject: '英',
          studentSubjects: { s2: '英', s1: '数' },
          regularSubstituteInfo: { s1: { regularTeacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' } },
        }],
      },
      availability: {
        'teacher:t1': ['2026-07-21_1'],
        'teacher:t2': ['2026-07-21_2'],
        'student:s1': ['2026-07-21_1', '2026-07-21_2'],
        'student:s2': ['2026-07-21_2'],
      },
    })
    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1', '2026-07-21_2', '2026-07-21_3'])
    const slot2 = result.assignments['2026-07-21_2'] ?? []
    expect(slot2).toHaveLength(1)
    expect(slot2[0].studentIds).toContain('s1')
    expect(slot2[0].regularSubstituteInfo?.s1).toEqual({ regularTeacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' })
  })

  it('keeps regular-like assignments for students without special submissions', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 2, holidays: [], lastAutoAssignedAt: Date.now() },
      teachers: [
        makeTeacher({ id: 't1', name: '通常講師', subjects: ['数'] }),
        makeTeacher({ id: 't2', name: '代行講師', subjects: ['数'] }),
      ],
      students: [
        makeStudent({ id: 's1', name: '通常生徒', subjects: [], subjectSlots: {}, submittedAt: 0 }),
      ],
      regularLessons: [
        { id: 'r1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 },
      ],
      assignments: {
        '2026-07-21_2': [{
          teacherId: 't2',
          studentIds: ['s1'],
          subject: '数',
          regularSubstituteInfo: { s1: { regularTeacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' } },
        }],
      },
      availability: {
        'teacher:t1': [],
        'teacher:t2': ['2026-07-21_2'],
        'student:s1': [],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1', '2026-07-21_2'])
    const slot2 = result.assignments['2026-07-21_2'] ?? []

    expect(slot2).toHaveLength(1)
    expect(slot2[0].studentIds).toEqual(['s1'])
    expect(slot2[0].regularSubstituteInfo?.s1).toEqual({ regularTeacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' })
  })

  it('fills an existing special single-student pair on rerun', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 1, holidays: [], lastAutoAssignedAt: Date.now() },
      teachers: [makeTeacher({ id: 't1', name: '田中先生', subjects: ['数'] })],
      students: [
        makeStudent({ id: 's1', name: 'A', subjects: ['数'], subjectSlots: { 数: 1 }, submittedAt: 1000 }),
        makeStudent({ id: 's2', name: 'B', subjects: ['数'], subjectSlots: { 数: 1 }, submittedAt: 2000 }),
      ],
      assignments: {
        '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数', studentSubjects: { s1: '数' } }],
      },
      availability: {
        'teacher:t1': ['2026-07-21_1'],
        'student:s1': ['2026-07-21_1'],
        'student:s2': ['2026-07-21_1'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1'])
    const slot1 = result.assignments['2026-07-21_1'] ?? []

    expect(slot1).toHaveLength(1)
    expect(slot1[0].teacherId).toBe('t1')
    expect(slot1[0].studentIds.sort()).toEqual(['s1', 's2'])
  })

  it('fills a regular single-student slot with a special-demand student on rerun', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 1, holidays: [], lastAutoAssignedAt: Date.now() },
      teachers: [makeTeacher({ id: 't1', name: '田中先生', subjects: ['数'] })],
      students: [
        makeStudent({ id: 's1', name: '通常生徒', subjects: [], subjectSlots: {}, submittedAt: 0 }),
        makeStudent({ id: 's2', name: '特別生徒', subjects: ['数'], subjectSlots: { 数: 1 }, submittedAt: 2000 }),
      ],
      regularLessons: [
        { id: 'r1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 },
      ],
      assignments: {
        '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数', studentSubjects: { s1: '数' }, isRegular: true }],
      },
      availability: {
        'teacher:t1': ['2026-07-21_1'],
        'student:s1': [],
        'student:s2': ['2026-07-21_1'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1'])
    const slot1 = result.assignments['2026-07-21_1'] ?? []

    expect(slot1).toHaveLength(1)
    expect(slot1[0].teacherId).toBe('t1')
    expect(slot1[0].studentIds.sort()).toEqual(['s1', 's2'])
    expect(slot1[0].studentSubjects?.s2 ?? slot1[0].subject).toBe('数')
  })

  it('assigns makeup for regular-lesson subjects even if the student special-subject list differs', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-22', slotsPerDay: 2, holidays: [] },
      teachers: [makeTeacher({ id: 't1', subjects: ['数', '英'] })],
      students: [makeStudent({
        id: 's1',
        subjects: ['英'],
        subjectSlots: { 英: 1 },
        unavailableSlots: ['2026-07-21_1'],
      })],
      regularLessons: [
        { id: 'r1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 },
      ],
      availability: {
        'teacher:t1': ['2026-07-21_1', '2026-07-22_1'],
        'student:s1': ['2026-07-22_1'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1', '2026-07-22_1'])
    const slotAssignments = result.assignments['2026-07-22_1'] ?? []
    const makeupAssignment = slotAssignments.find((assignment) => assignment.regularMakeupInfo?.s1)

    expect(makeupAssignment).toBeDefined()
    expect(makeupAssignment?.teacherId).toBe('t1')
    expect(makeupAssignment?.studentIds).toContain('s1')
    expect(makeupAssignment?.regularMakeupInfo?.s1).toEqual({ dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' })
    expect(makeupAssignment?.studentSubjects?.s1 ?? makeupAssignment?.subject).toBe('数')
  })

  it('includes makeup-only students in empty-slot candidate selection', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-22', slotsPerDay: 2, holidays: [] },
      teachers: [makeTeacher({ id: 't1', subjects: ['数'] })],
      students: [makeStudent({
        id: 's1',
        subjects: ['英'],
        subjectSlots: { 英: 1 },
        unavailableSlots: ['2026-07-21_1', '2026-07-21_2'],
      })],
      regularLessons: [
        { id: 'r1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 },
      ],
      availability: {
        'teacher:t1': ['2026-07-21_1', '2026-07-22_1'],
        'student:s1': ['2026-07-22_1'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1', '2026-07-22_1'])
    const slotAssignments = result.assignments['2026-07-22_1'] ?? []
    const makeupAssignment = slotAssignments.find((assignment) => assignment.regularMakeupInfo?.s1)

    expect(makeupAssignment).toBeDefined()
    expect(makeupAssignment?.teacherId).toBe('t1')
    expect(makeupAssignment?.studentIds).toContain('s1')
    expect(makeupAssignment?.studentSubjects?.s1 ?? makeupAssignment?.subject).toBe('数')
  })

  it('removes duplicate makeup coverage for the same regular occurrence on rerun', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-22', slotsPerDay: 2, holidays: [], lastAutoAssignedAt: Date.now() },
      regularLessons: [
        { id: 'r1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 },
      ],
      assignments: {
        '2026-07-21_2': [{
          teacherId: 't1',
          studentIds: ['s1'],
          subject: '数',
          regularMakeupInfo: { s1: { dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' } },
        }],
        '2026-07-22_1': [{
          teacherId: 't1',
          studentIds: ['s1'],
          subject: '数',
          regularMakeupInfo: { s1: { dayOfWeek: 2, slotNumber: 1, date: '2026-07-21' } },
        }],
      },
      availability: {
        'teacher:t1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
        'student:s1': ['2026-07-21_1', '2026-07-21_2', '2026-07-22_1', '2026-07-22_2'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, slots)
    const remainingSlots = Object.entries(result.assignments)
      .flatMap(([slot, assignments]) => assignments.filter((assignment) => assignment.regularMakeupInfo?.s1).map(() => slot))

    expect(remainingSlots).toEqual(['2026-07-21_2'])
    expect((result.assignments['2026-07-22_1'] ?? []).every((assignment) => !assignment.regularMakeupInfo?.s1)).toBe(true)
  })

  it('deduplicates a makeup demand recreated by actual absence', () => {
    const deduped = dedupeMakeupInfosByOccurrence('s1', [
      { teacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21', subject: '数' },
      { teacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21', subject: '数', absentDate: '2026-07-21', reasonKind: 'actual-absence' as const },
    ])

    expect(deduped).toEqual([
      { teacherId: 't1', dayOfWeek: 2, slotNumber: 1, date: '2026-07-21', subject: '数', absentDate: '2026-07-21', reasonKind: 'actual-absence' },
    ])
  })

  it('prefers 2-student pairs', async () => {
    const data = makeSessionData({
      teachers: [makeTeacher()],
      students: [
        makeStudent({ id: 's1', name: 'A', subjects: ['数'], subjectSlots: { '数': 1 }, submittedAt: 1000 }),
        makeStudent({ id: 's2', name: 'B', subjects: ['数'], subjectSlots: { '数': 1 }, submittedAt: 2000 }),
      ],
      availability: {
        'teacher:t1': ['2026-07-21_1'],
        'student:s1': ['2026-07-21_1'],
        'student:s2': ['2026-07-21_1'],
      },
    })
    const result = await buildIncrementalAutoAssignments(data, slots)
    const slot1 = result.assignments['2026-07-21_1'] ?? []
    // Should pair both students together with the teacher
    const pairAssign = slot1.find(a => a.teacherId === 't1')
    expect(pairAssign?.studentIds).toHaveLength(2)
  })

  it('respects desk count limit', async () => {
    const data = makeSessionData({
      settings: { name: '', adminPassword: '', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 1, holidays: [], deskCount: 1 },
      teachers: [
        makeTeacher({ id: 't1', name: 'A先生' }),
        makeTeacher({ id: 't2', name: 'B先生' }),
      ],
      students: [
        makeStudent({ id: 's1', subjects: ['数'], subjectSlots: { '数': 1 }, submittedAt: 1000 }),
        makeStudent({ id: 's2', subjects: ['数'], subjectSlots: { '数': 1 }, submittedAt: 2000 }),
      ],
      availability: {
        'teacher:t1': ['2026-07-21_1'],
        'teacher:t2': ['2026-07-21_1'],
        'student:s1': ['2026-07-21_1'],
        'student:s2': ['2026-07-21_1'],
      },
    })
    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1'])
    const slot1 = result.assignments['2026-07-21_1'] ?? []
    expect(slot1.length).toBeLessThanOrEqual(1)
  })

  it('supplementary assignment still uses an otherwise valid slot even if the same teacher taught the student in the previous slot', async () => {
    const data = makeSessionData({
      settings: { name: '春期講習', adminPassword: 'admin', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 2, holidays: [], lastAutoAssignedAt: Date.now() },
      teachers: [makeTeacher({ id: 't1', name: '田中先生', subjects: ['数'] })],
      students: [
        makeStudent({ id: 's1', name: '先行生徒', subjects: ['数'], subjectSlots: { 数: 1 }, submittedAt: 1000 }),
        makeStudent({ id: 's2', name: '残コマ生徒', subjects: ['数'], subjectSlots: { 数: 2 }, submittedAt: 2000 }),
      ],
      assignments: {
        '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s2'], subject: '数', studentSubjects: { s2: '数' } }],
      },
      availability: {
        'teacher:t1': ['2026-07-21_1', '2026-07-21_2'],
        'student:s1': ['2026-07-21_1'],
        'student:s2': ['2026-07-21_1', '2026-07-21_2'],
      },
    })

    const result = await buildIncrementalAutoAssignments(data, ['2026-07-21_1', '2026-07-21_2'])
    const slot2 = result.assignments['2026-07-21_2'] ?? []

    expect(slot2.some((assignment) => assignment.teacherId === 't1' && assignment.studentIds.includes('s2'))).toBe(true)
  })
})

describe('buildMendanAutoAssignments', () => {
  const slots = ['2026-07-21_1', '2026-07-21_2']

  it('assigns parents to managers FCFS', () => {
    const data = makeSessionData({
      settings: { name: '', adminPassword: '', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 2, holidays: [], sessionType: 'mendan' },
      managers: [{ id: 'm1', name: 'マネ', email: '' }],
      students: [
        makeStudent({ id: 'p1', name: '保護者A', submittedAt: 1000 }),
        makeStudent({ id: 'p2', name: '保護者B', submittedAt: 2000 }),
      ],
      availability: {
        'manager:m1': ['2026-07-21_1', '2026-07-21_2'],
        'student:p1': ['2026-07-21_1', '2026-07-21_2'],
        'student:p2': ['2026-07-21_1', '2026-07-21_2'],
      },
      assignments: {},
    })
    const result = buildMendanAutoAssignments(data, slots)
    const allAssigns = Object.values(result.assignments).flat()
    expect(allAssigns.filter(a => a.studentIds.includes('p1'))).toHaveLength(1)
    expect(allAssigns.filter(a => a.studentIds.includes('p2'))).toHaveLength(1)
    expect(result.unassignedParents).toHaveLength(0)
  })

  it('skips unsubmitted parents', () => {
    const data = makeSessionData({
      managers: [{ id: 'm1', name: 'マネ', email: '' }],
      students: [makeStudent({ id: 'p1', name: '保護者A', submittedAt: 0 })],
      availability: {
        'manager:m1': ['2026-07-21_1'],
        'student:p1': ['2026-07-21_1'],
      },
      assignments: {},
    })
    const result = buildMendanAutoAssignments(data, slots)
    const allAssigns = Object.values(result.assignments).flat()
    expect(allAssigns).toHaveLength(0)
  })

  it('reports unassigned parents when no slots available', () => {
    const data = makeSessionData({
      managers: [{ id: 'm1', name: 'マネ', email: '' }],
      students: [makeStudent({ id: 'p1', name: '保護者A', submittedAt: 1000 })],
      availability: {
        'manager:m1': [], // manager has no availability
        'student:p1': ['2026-07-21_1'],
      },
      assignments: {},
    })
    const result = buildMendanAutoAssignments(data, slots)
    expect(result.unassignedParents).toContain('保護者A')
  })

  it('respects desk count limit', () => {
    const data = makeSessionData({
      settings: { name: '', adminPassword: '', startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 1, holidays: [], deskCount: 1 },
      managers: [
        { id: 'm1', name: 'マネA', email: '' },
        { id: 'm2', name: 'マネB', email: '' },
      ],
      students: [
        makeStudent({ id: 'p1', name: '保護者A', submittedAt: 1000 }),
        makeStudent({ id: 'p2', name: '保護者B', submittedAt: 2000 }),
      ],
      availability: {
        'manager:m1': ['2026-07-21_1'],
        'manager:m2': ['2026-07-21_1'],
        'student:p1': ['2026-07-21_1'],
        'student:p2': ['2026-07-21_1'],
      },
      assignments: {},
    })
    const result = buildMendanAutoAssignments(data, ['2026-07-21_1'])
    const slot1 = result.assignments['2026-07-21_1'] ?? []
    expect(slot1.length).toBeLessThanOrEqual(1)
  })
})
