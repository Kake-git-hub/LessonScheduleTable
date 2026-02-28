import { constraintFor, hasAvailability, isStudentAvailable, isParentAvailableForMendan, isRegularLessonPair } from '../constraints'
import type { PairConstraint, Student, RegularLesson } from '../../types'

describe('constraintFor', () => {
  const constraints: PairConstraint[] = [
    { id: '1', personAId: 'a', personBId: 'b', personAType: 'teacher', personBType: 'student', type: 'incompatible' },
    { id: '2', personAId: 'c', personBId: 'd', personAType: 'student', personBType: 'student', type: 'recommended' },
  ]

  it('returns null when no matching constraint exists', () => {
    expect(constraintFor(constraints, 'x', 'y')).toBeNull()
  })

  it('returns incompatible when A-B pair matches', () => {
    expect(constraintFor(constraints, 'a', 'b')).toBe('incompatible')
  })

  it('returns incompatible when B-A pair matches (order-independent)', () => {
    expect(constraintFor(constraints, 'b', 'a')).toBe('incompatible')
  })

  it('returns recommended for recommended constraints', () => {
    expect(constraintFor(constraints, 'c', 'd')).toBe('recommended')
  })

  it('handles empty constraints array', () => {
    expect(constraintFor([], 'a', 'b')).toBeNull()
  })
})

describe('hasAvailability', () => {
  const availability = {
    'teacher:t1': ['2026-07-21_1', '2026-07-21_2'],
    'student:s1': ['2026-07-21_1'],
  }

  it('returns true when person has availability for the slot', () => {
    expect(hasAvailability(availability, 'teacher', 't1', '2026-07-21_1')).toBe(true)
  })

  it('returns false when slot is not in availability array', () => {
    expect(hasAvailability(availability, 'teacher', 't1', '2026-07-21_3')).toBe(false)
  })

  it('returns false when person has no availability entry', () => {
    expect(hasAvailability(availability, 'teacher', 't99', '2026-07-21_1')).toBe(false)
  })
})

describe('isStudentAvailable', () => {
  const makeStudent = (overrides: Partial<Student> = {}): Student => ({
    id: 's1',
    name: 'Test',
    email: '',
    grade: '中1',
    subjects: ['数'],
    subjectSlots: { '数': 3 },
    unavailableDates: [],
    preferredSlots: [],
    unavailableSlots: [],
    memo: '',
    submittedAt: 1000,
    ...overrides,
  })

  it('returns false for unsubmitted student (submittedAt === 0)', () => {
    expect(isStudentAvailable(makeStudent({ submittedAt: 0 }), '2026-07-21_1')).toBe(false)
  })

  it('returns false when slot is in unavailableSlots', () => {
    expect(isStudentAvailable(makeStudent({ unavailableSlots: ['2026-07-21_1'] }), '2026-07-21_1')).toBe(false)
  })

  it('returns false when date is in unavailableDates (legacy)', () => {
    expect(isStudentAvailable(makeStudent({ unavailableDates: ['2026-07-21'] }), '2026-07-21_1')).toBe(false)
  })

  it('returns true when student is available', () => {
    expect(isStudentAvailable(makeStudent(), '2026-07-21_1')).toBe(true)
  })
})

describe('isParentAvailableForMendan', () => {
  it('returns true when parent has availability for the slot', () => {
    const availability = { 'student:s1': ['2026-07-21_1'] }
    expect(isParentAvailableForMendan(availability, 's1', '2026-07-21_1')).toBe(true)
  })

  it('returns false when slot is not in availability', () => {
    const availability = { 'student:s1': ['2026-07-21_1'] }
    expect(isParentAvailableForMendan(availability, 's1', '2026-07-21_2')).toBe(false)
  })
})

describe('isRegularLessonPair', () => {
  const regularLessons: RegularLesson[] = [
    { id: '1', teacherId: 't1', studentIds: ['s1', 's2'], subject: '数', dayOfWeek: 1, slotNumber: 1 },
  ]

  it('returns true when teacher-student pair exists', () => {
    expect(isRegularLessonPair(regularLessons, 't1', 's1')).toBe(true)
  })

  it('returns true for second student in pair', () => {
    expect(isRegularLessonPair(regularLessons, 't1', 's2')).toBe(true)
  })

  it('returns false when no matching pair', () => {
    expect(isRegularLessonPair(regularLessons, 't1', 's99')).toBe(false)
  })

  it('returns false for wrong teacher', () => {
    expect(isRegularLessonPair(regularLessons, 't99', 's1')).toBe(false)
  })
})
