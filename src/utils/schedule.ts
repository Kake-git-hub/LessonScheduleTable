import type { SessionData } from '../types'

const toIso = (date: Date): string => {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}

export const buildSlotKeys = (settings: SessionData['settings']): string[] => {
  if (!settings.startDate || !settings.endDate || settings.slotsPerDay <= 0) {
    return []
  }

  const start = new Date(`${settings.startDate}T00:00:00`)
  const end = new Date(`${settings.endDate}T00:00:00`)
  const holidaySet = new Set(settings.holidays)
  const result: string[] = []

  for (let cursor = new Date(start); cursor <= end; cursor.setDate(cursor.getDate() + 1)) {
    const iso = toIso(cursor)
    if (holidaySet.has(iso)) {
      continue
    }

    for (let slot = 1; slot <= settings.slotsPerDay; slot += 1) {
      result.push(`${iso}_${slot}`)
    }
  }

  return result
}

export const slotLabel = (slotKey: string): string => {
  const [date, slot] = slotKey.split('_')
  return `${date} ${slot}限`
}

export const personKey = (personType: 'teacher' | 'student', personId: string): string =>
  `${personType}:${personId}`
