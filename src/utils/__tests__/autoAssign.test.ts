import { buildIncrementalAutoAssignments, buildMendanAutoAssignments } from '../autoAssign'
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
