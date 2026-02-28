import {
  getSlotNumber, getIsoDayOfWeek, getSlotDayOfWeek,
  allAssignments, buildEffectiveAssignments, getStudentSubject,
  countTeacherLoad, getTeacherAssignedDates, getTeacherSlotNumbersOnDate,
  getTeacherPrevSlotStudentIds,
  countStudentSlotsOnDate, getStudentSlotNumbersOnDate, countStudentAssignedDates,
  countStudentLoad, countStudentSubjectLoad,
  collectTeacherShortages, assignmentSignature, hasMeaningfulManualAssignment,
  getTeacherStudentSlotsOnDate, getStudentSubjectsOnAdjacentSlots,
  findRegularLessonsForSlot, getDatesInRange,
} from '../assignments'
import type { Assignment, SessionData } from '../../types'

// Helper to create minimal SessionData for collectTeacherShortages tests
const makeSessionData = (overrides: Partial<SessionData> = {}): SessionData => ({
  settings: { name: '', adminPassword: '', startDate: '2026-07-21', endDate: '2026-07-23', slotsPerDay: 3, holidays: [] },
  subjects: ['数', '英'],
  managers: [],
  teachers: [{ id: 't1', name: '田中', email: '', subjects: ['数', '英'], memo: '' }],
  students: [{ id: 's1', name: '山田', email: '', grade: '中1', subjects: ['数'], subjectSlots: { '数': 3 }, unavailableDates: [], preferredSlots: [], unavailableSlots: [], memo: '', submittedAt: 1000 }],
  constraints: [],
  gradeConstraints: [],
  availability: { 'teacher:t1': ['2026-07-21_1', '2026-07-21_2'] },
  assignments: {},
  regularLessons: [],
  ...overrides,
})

describe('getSlotNumber', () => {
  it('parses slot number from key', () => {
    expect(getSlotNumber('2026-07-21_3')).toBe(3)
  })

  it('handles single digit', () => {
    expect(getSlotNumber('2026-07-21_1')).toBe(1)
  })
})

describe('getIsoDayOfWeek', () => {
  it('returns 0 for Sunday', () => {
    // 2026-07-19 is Sunday
    expect(getIsoDayOfWeek('2026-07-19')).toBe(0)
  })

  it('returns 2 for Tuesday', () => {
    // 2026-07-21 is Tuesday
    expect(getIsoDayOfWeek('2026-07-21')).toBe(2)
  })
})

describe('getSlotDayOfWeek', () => {
  it('extracts date and returns day of week', () => {
    expect(getSlotDayOfWeek('2026-07-21_3')).toBe(2) // Tuesday
  })
})

describe('allAssignments', () => {
  it('flattens all slot arrays', () => {
    const assignments: Record<string, Assignment[]> = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_2': [{ teacherId: 't2', studentIds: ['s2'], subject: '英' }],
    }
    expect(allAssignments(assignments)).toHaveLength(2)
  })

  it('returns empty array for empty input', () => {
    expect(allAssignments({})).toEqual([])
  })
})

describe('buildEffectiveAssignments', () => {
  it('returns original assignments when no actualResults', () => {
    const assignments = { '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }] }
    const result = buildEffectiveAssignments(assignments)
    expect(result['2026-07-21_1']).toHaveLength(1)
    expect(result['2026-07-21_1'][0].teacherId).toBe('t1')
  })

  it('overrides specific slots with actual results', () => {
    const assignments = { '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }] }
    const actualResults = { '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s2'], subject: '英' }] }
    const result = buildEffectiveAssignments(assignments, actualResults)
    expect(result['2026-07-21_1'][0].studentIds).toEqual(['s2'])
    expect(result['2026-07-21_1'][0].subject).toBe('英')
  })
})

describe('getStudentSubject', () => {
  it('returns assignment subject when no studentSubjects', () => {
    const a: Assignment = { teacherId: 't1', studentIds: ['s1'], subject: '数' }
    expect(getStudentSubject(a, 's1')).toBe('数')
  })

  it('returns per-student subject when set', () => {
    const a: Assignment = { teacherId: 't1', studentIds: ['s1', 's2'], subject: '数', studentSubjects: { s1: '英', s2: '数' } }
    expect(getStudentSubject(a, 's1')).toBe('英')
    expect(getStudentSubject(a, 's2')).toBe('数')
  })
})

describe('countTeacherLoad', () => {
  it('counts total assignments for a teacher', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_2': [{ teacherId: 't1', studentIds: ['s2'], subject: '英' }, { teacherId: 't2', studentIds: ['s3'], subject: '数' }],
    }
    expect(countTeacherLoad(assignments, 't1')).toBe(2)
    expect(countTeacherLoad(assignments, 't2')).toBe(1)
  })

  it('returns 0 for unassigned teacher', () => {
    expect(countTeacherLoad({}, 't99')).toBe(0)
  })
})

describe('getTeacherAssignedDates', () => {
  it('returns set of unique dates', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_2': [{ teacherId: 't1', studentIds: ['s2'], subject: '英' }],
      '2026-07-22_1': [{ teacherId: 't1', studentIds: ['s3'], subject: '数' }],
    }
    const dates = getTeacherAssignedDates(assignments, 't1')
    expect(dates.size).toBe(2)
    expect(dates.has('2026-07-21')).toBe(true)
    expect(dates.has('2026-07-22')).toBe(true)
  })
})

describe('getTeacherSlotNumbersOnDate', () => {
  it('returns sorted slot numbers', () => {
    const assignments = {
      '2026-07-21_3': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s2'], subject: '英' }],
    }
    expect(getTeacherSlotNumbersOnDate(assignments, 't1', '2026-07-21')).toEqual([1, 3])
  })
})

describe('getTeacherPrevSlotStudentIds', () => {
  it('returns student IDs from previous slot', () => {
    const assignments = {
      '2026-07-21_2': [{ teacherId: 't1', studentIds: ['s1', 's2'], subject: '数' }],
    }
    expect(getTeacherPrevSlotStudentIds(assignments, 't1', '2026-07-21', 3)).toEqual(['s1', 's2'])
  })

  it('returns empty array when no previous slot', () => {
    expect(getTeacherPrevSlotStudentIds({}, 't1', '2026-07-21', 1)).toEqual([])
  })
})

describe('countStudentSlotsOnDate', () => {
  it('counts all assignments on a specific date', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_2': [{ teacherId: 't2', studentIds: ['s1'], subject: '英' }],
      '2026-07-22_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
    }
    expect(countStudentSlotsOnDate(assignments, 's1', '2026-07-21')).toBe(2)
  })
})

describe('getStudentSlotNumbersOnDate', () => {
  it('returns sorted slot numbers for student', () => {
    const assignments = {
      '2026-07-21_3': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_1': [{ teacherId: 't2', studentIds: ['s1'], subject: '英' }],
    }
    expect(getStudentSlotNumbersOnDate(assignments, 's1', '2026-07-21')).toEqual([1, 3])
  })
})

describe('countStudentAssignedDates', () => {
  it('returns number of unique dates', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-22_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-22_2': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
    }
    expect(countStudentAssignedDates(assignments, 's1')).toBe(2)
  })
})

describe('countStudentLoad', () => {
  it('counts non-regular assignments only', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_2': [{ teacherId: 't1', studentIds: ['s1'], subject: '数', isRegular: true }],
    }
    expect(countStudentLoad(assignments, 's1')).toBe(1)
  })
})

describe('countStudentSubjectLoad', () => {
  it('counts per-subject non-regular assignments', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_2': [{ teacherId: 't1', studentIds: ['s1'], subject: '英' }],
      '2026-07-21_3': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
    }
    expect(countStudentSubjectLoad(assignments, 's1', '数')).toBe(2)
    expect(countStudentSubjectLoad(assignments, 's1', '英')).toBe(1)
  })
})

describe('collectTeacherShortages', () => {
  it('detects missing teacher (empty teacherId)', () => {
    const data = makeSessionData()
    const assignments = { '2026-07-21_1': [{ teacherId: '', studentIds: ['s1'], subject: '数' }] }
    const result = collectTeacherShortages(data, assignments)
    expect(result).toHaveLength(1)
    expect(result[0].detail).toBe('講師未設定')
  })

  it('detects deleted teacher', () => {
    const data = makeSessionData()
    const assignments = { '2026-07-21_1': [{ teacherId: 't99', studentIds: ['s1'], subject: '数' }] }
    const result = collectTeacherShortages(data, assignments)
    expect(result).toHaveLength(1)
    expect(result[0].detail).toContain('未登録')
  })

  it('detects availability conflict', () => {
    const data = makeSessionData()
    const assignments = { '2026-07-21_3': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }] }
    // t1 is only available for slots 1 and 2
    const result = collectTeacherShortages(data, assignments)
    expect(result).toHaveLength(1)
    expect(result[0].detail).toContain('出席不可')
  })

  it('detects subject mismatch', () => {
    const data = makeSessionData()
    const assignments = { '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '理' }] }
    const result = collectTeacherShortages(data, assignments)
    expect(result).toHaveLength(1)
    expect(result[0].detail).toContain('担当外科目')
  })

  it('skips regular assignments', () => {
    const data = makeSessionData()
    const assignments = { '2026-07-21_1': [{ teacherId: 't99', studentIds: ['s1'], subject: '数', isRegular: true }] }
    const result = collectTeacherShortages(data, assignments)
    expect(result).toHaveLength(0)
  })
})

describe('assignmentSignature', () => {
  it('produces deterministic signature', () => {
    const a: Assignment = { teacherId: 't1', studentIds: ['s2', 's1'], subject: '数' }
    const sig = assignmentSignature(a)
    expect(sig).toContain('t1')
    expect(sig).toContain('s1')
    expect(sig).toContain('s2')
  })

  it('differentiates regular vs non-regular', () => {
    const a1: Assignment = { teacherId: 't1', studentIds: ['s1'], subject: '数' }
    const a2: Assignment = { teacherId: 't1', studentIds: ['s1'], subject: '数', isRegular: true }
    expect(assignmentSignature(a1)).not.toBe(assignmentSignature(a2))
  })
})

describe('hasMeaningfulManualAssignment', () => {
  it('returns true for assignment with teacher', () => {
    expect(hasMeaningfulManualAssignment({ teacherId: 't1', studentIds: [], subject: '' })).toBe(true)
  })

  it('returns false for regular assignment', () => {
    expect(hasMeaningfulManualAssignment({ teacherId: 't1', studentIds: ['s1'], subject: '数', isRegular: true })).toBe(false)
  })

  it('returns false for empty non-regular assignment', () => {
    expect(hasMeaningfulManualAssignment({ teacherId: '', studentIds: [], subject: '' })).toBe(false)
  })
})

describe('getTeacherStudentSlotsOnDate', () => {
  it('returns slot numbers for a specific teacher-student pair', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_3': [{ teacherId: 't1', studentIds: ['s1'], subject: '英' }],
      '2026-07-21_2': [{ teacherId: 't1', studentIds: ['s2'], subject: '数' }],
    }
    expect(getTeacherStudentSlotsOnDate(assignments, 't1', 's1', '2026-07-21')).toEqual([1, 3])
  })
})

describe('getStudentSubjectsOnAdjacentSlots', () => {
  it('returns subjects from adjacent slots', () => {
    const assignments = {
      '2026-07-21_1': [{ teacherId: 't1', studentIds: ['s1'], subject: '数' }],
      '2026-07-21_3': [{ teacherId: 't2', studentIds: ['s1'], subject: '英' }],
    }
    expect(getStudentSubjectsOnAdjacentSlots(assignments, 's1', '2026-07-21', 2)).toEqual(['数', '英'])
  })

  it('returns empty when no adjacent assignments', () => {
    expect(getStudentSubjectsOnAdjacentSlots({}, 's1', '2026-07-21', 2)).toEqual([])
  })
})

describe('findRegularLessonsForSlot', () => {
  const regularLessons = [
    { id: '1', teacherId: 't1', studentIds: ['s1'], subject: '数', dayOfWeek: 2, slotNumber: 1 }, // Tuesday slot 1
    { id: '2', teacherId: 't2', studentIds: ['s2'], subject: '英', dayOfWeek: 2, slotNumber: 2 },
  ]

  it('returns matching regular lessons by day and slot', () => {
    // 2026-07-21 is Tuesday (dayOfWeek=2)
    const result = findRegularLessonsForSlot(regularLessons, '2026-07-21_1')
    expect(result).toHaveLength(1)
    expect(result[0].teacherId).toBe('t1')
  })

  it('returns empty for non-matching slot', () => {
    const result = findRegularLessonsForSlot(regularLessons, '2026-07-21_3')
    expect(result).toHaveLength(0)
  })
})

describe('getDatesInRange', () => {
  it('generates correct dates between start and end', () => {
    const settings = { name: '', adminPassword: '', startDate: '2026-07-21', endDate: '2026-07-23', slotsPerDay: 3, holidays: [] as string[] }
    expect(getDatesInRange(settings)).toEqual(['2026-07-21', '2026-07-22', '2026-07-23'])
  })

  it('excludes holidays', () => {
    const settings = { name: '', adminPassword: '', startDate: '2026-07-21', endDate: '2026-07-23', slotsPerDay: 3, holidays: ['2026-07-22'] }
    expect(getDatesInRange(settings)).toEqual(['2026-07-21', '2026-07-23'])
  })

  it('returns empty for missing dates', () => {
    const settings = { name: '', adminPassword: '', startDate: '', endDate: '', slotsPerDay: 3, holidays: [] as string[] }
    expect(getDatesInRange(settings)).toEqual([])
  })
})
