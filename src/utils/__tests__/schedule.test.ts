import { buildSlotKeys, formatShortDate, slotLabel, mendanTimeLabel, personKey } from '../schedule'

describe('buildSlotKeys', () => {
  it('returns empty array when startDate is missing', () => {
    const settings = { startDate: '', endDate: '2026-07-25', slotsPerDay: 3, holidays: [] as string[], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual([])
  })

  it('returns empty array when endDate is missing', () => {
    const settings = { startDate: '2026-07-21', endDate: '', slotsPerDay: 3, holidays: [] as string[], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual([])
  })

  it('returns empty array when slotsPerDay is 0', () => {
    const settings = { startDate: '2026-07-21', endDate: '2026-07-22', slotsPerDay: 0, holidays: [] as string[], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual([])
  })

  it('generates correct keys for a 2-day range with 3 slots/day', () => {
    const settings = { startDate: '2026-07-21', endDate: '2026-07-22', slotsPerDay: 3, holidays: [] as string[], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual([
      '2026-07-21_1', '2026-07-21_2', '2026-07-21_3',
      '2026-07-22_1', '2026-07-22_2', '2026-07-22_3',
    ])
  })

  it('excludes holidays', () => {
    const settings = { startDate: '2026-07-21', endDate: '2026-07-23', slotsPerDay: 2, holidays: ['2026-07-22'], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual([
      '2026-07-21_1', '2026-07-21_2',
      '2026-07-23_1', '2026-07-23_2',
    ])
  })

  it('handles single-day range', () => {
    const settings = { startDate: '2026-07-21', endDate: '2026-07-21', slotsPerDay: 1, holidays: [] as string[], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual(['2026-07-21_1'])
  })

  it('returns empty when all days are holidays', () => {
    const settings = { startDate: '2026-07-21', endDate: '2026-07-22', slotsPerDay: 2, holidays: ['2026-07-21', '2026-07-22'], name: '', adminPassword: '' }
    expect(buildSlotKeys(settings)).toEqual([])
  })
})

describe('formatShortDate', () => {
  it('formats a mid-year date correctly', () => {
    // 2026-07-21 is a Tuesday (火)
    expect(formatShortDate('2026-07-21')).toBe('7/21(火)')
  })

  it('formats January 1st correctly', () => {
    // 2026-01-01 is a Thursday (木)
    expect(formatShortDate('2026-01-01')).toBe('1/1(木)')
  })

  it('formats December 31st correctly', () => {
    // 2026-12-31 is a Thursday (木)
    expect(formatShortDate('2026-12-31')).toBe('12/31(木)')
  })

  it('shows correct day-of-week kanji for Sunday', () => {
    // 2026-07-19 is a Sunday (日)
    expect(formatShortDate('2026-07-19')).toBe('7/19(日)')
  })
})

describe('slotLabel', () => {
  it('returns "M/D(曜) N限" for normal mode', () => {
    expect(slotLabel('2026-07-21_3')).toBe('7/21(火) 3限')
  })

  it('returns time format for mendan mode with default startHour', () => {
    // slot 2, startHour=10 → hour = 10 - 1 + 2 = 11
    expect(slotLabel('2026-07-21_2', true)).toBe('7/21(火) 11:00')
  })

  it('returns time format for mendan mode with custom startHour', () => {
    // slot 3, startHour=14 → hour = 14 - 1 + 3 = 16
    expect(slotLabel('2026-07-21_3', true, 14)).toBe('7/21(火) 16:00')
  })
})

describe('mendanTimeLabel', () => {
  it('returns correct time with default startHour', () => {
    // slot 1, startHour=10 → 10 - 1 + 1 = 10
    expect(mendanTimeLabel(1)).toBe('10:00')
  })

  it('returns correct time for slot 2 with default startHour', () => {
    // slot 2, startHour=10 → 10 - 1 + 2 = 11
    expect(mendanTimeLabel(2)).toBe('11:00')
  })

  it('returns correct time with custom startHour', () => {
    // slot 3, startHour=14 → 14 - 1 + 3 = 16
    expect(mendanTimeLabel(3, 14)).toBe('16:00')
  })
})

describe('personKey', () => {
  it('returns "teacher:abc" for teacher type', () => {
    expect(personKey('teacher', 'abc')).toBe('teacher:abc')
  })

  it('returns "student:xyz" for student type', () => {
    expect(personKey('student', 'xyz')).toBe('student:xyz')
  })

  it('returns "manager:m1" for manager type', () => {
    expect(personKey('manager', 'm1')).toBe('manager:m1')
  })
})
