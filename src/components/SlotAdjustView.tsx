import { useMemo, useState, useEffect, useCallback, useRef } from 'react'
import type { SessionData, Teacher } from '../types'
import { getSlotNumber, getIsoDayOfWeek, getStudentSubject } from '../utils/assignments'
import { hasAvailability, isStudentAvailable } from '../utils/constraints'

// ---- types ----

type SlotAdjustViewProps = {
  data: SessionData
  instructors: Teacher[]
  instructorPersonType: 'teacher' | 'manager'
  isMendan: boolean
  onMove: (
    sourceSlot: string,
    sourceIdx: number,
    studentId: string,
    targetSlot: string,
    targetIdx?: number,
  ) => Promise<void>
  onClose: () => void
}

type SelectionInfo = {
  slot: string
  assignmentIdx: number
  studentId: string
  studentName: string
}

// ---- helpers ----

const DAY_NAMES = ['日', '月', '火', '水', '木', '金', '土']

function getAllDatesInRange(startDate: string, endDate: string): string[] {
  if (!startDate || !endDate) return []
  const start = new Date(`${startDate}T00:00:00`)
  const end = new Date(`${endDate}T00:00:00`)
  const result: string[] = []
  for (let cur = new Date(start); cur <= end; cur = new Date(cur.getTime() + 86400000)) {
    const y = cur.getFullYear()
    const m = String(cur.getMonth() + 1).padStart(2, '0')
    const d = String(cur.getDate()).padStart(2, '0')
    result.push(`${y}-${m}-${d}`)
  }
  return result
}

function splitIntoWeeks(dates: string[]): string[][] {
  if (dates.length === 0) return []
  const weeks: string[][] = []
  let current: string[] = []
  for (const date of dates) {
    const dow = getIsoDayOfWeek(date)
    if (dow === 1 && current.length > 0) {
      weeks.push(current)
      current = []
    }
    current.push(date)
  }
  if (current.length > 0) weeks.push(current)
  return weeks
}

function padWeek(weekDates: string[]): string[] {
  const firstDate = weekDates[0]
  const full: string[] = []
  const firstDow = getIsoDayOfWeek(firstDate)
  const startPad = firstDow === 0 ? 6 : firstDow - 1
  const firstMs = new Date(`${firstDate}T00:00:00`).getTime()
  for (let p = startPad; p > 0; p--) {
    const d = new Date(firstMs - p * 86400000)
    full.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`)
  }
  for (const date of weekDates) full.push(date)
  const lastMs = new Date(`${weekDates[weekDates.length - 1]}T00:00:00`).getTime()
  let tailIdx = 1
  while (full.length < 7) {
    const d = new Date(lastMs + tailIdx * 86400000)
    full.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`)
    tailIdx++
  }
  return full
}

const SLOT_TIME_LABELS = [
  '13:00～14:30',
  '14:40～16:10',
  '16:20～17:50',
  '18:00～19:30',
  '19:40～21:10',
]

function formatSlotTime(slotNumber: number): string {
  return SLOT_TIME_LABELS[slotNumber - 1] ?? `${slotNumber}限`
}

function hasInstructorAvailabilityLocal(
  data: SessionData,
  isMendan: boolean,
  personType: 'teacher' | 'manager',
  id: string,
  slot: string,
): boolean {
  if (hasAvailability(data.availability, personType, id, slot)) return true
  if (personType !== 'teacher' || isMendan) return false
  if ((data.teacherSubmittedAt?.[id] ?? 0) > 0) return false
  const [date] = slot.split('_')
  const dayOfWeek = getIsoDayOfWeek(date)
  const slotNumber = getSlotNumber(slot)
  return data.regularLessons.some(
    (lesson) => lesson.teacherId === id && lesson.dayOfWeek === dayOfWeek && lesson.slotNumber === slotNumber,
  )
}

// ---- component ----

export default function SlotAdjustView({
  data,
  instructors,
  instructorPersonType,
  isMendan,
  onMove,
  onClose,
}: SlotAdjustViewProps) {
  const [selection, setSelection] = useState<SelectionInfo | null>(null)
  const [weekIdx, setWeekIdx] = useState(0)
  const [moving, setMoving] = useState(false)
  const containerRef = useRef<HTMLDivElement>(null)

  const holidays = useMemo(() => new Set(data.settings.holidays), [data.settings.holidays])
  const allDates = useMemo(
    () => getAllDatesInRange(data.settings.startDate, data.settings.endDate),
    [data.settings.startDate, data.settings.endDate],
  )
  const weeks = useMemo(() => splitIntoWeeks(allDates), [allDates])
  const lectureDateSet = useMemo(() => new Set(allDates), [allDates])

  const slotsPerDay = data.settings.slotsPerDay
  const deskCount = data.settings.deskCount && data.settings.deskCount > 0
    ? data.settings.deskCount
    : Math.max(1, ...Object.values(data.assignments).map((a) => a.length))

  // Student availability check
  const isStudentAvailableForSlot = useCallback(
    (studentId: string, slot: string): boolean => {
      const student = data.students.find((s) => s.id === studentId)
      if (!student) return false
      if (isMendan) {
        // For mendan, check parent availability
        const key = `student:${studentId}`
        return (data.availability[key] ?? []).includes(slot)
      }
      return isStudentAvailable(student, slot)
    },
    [data, isMendan],
  )

  // Check if a student is already in this slot
  const isStudentInSlot = useCallback(
    (studentId: string, slot: string, excludeSlot?: string, excludeIdx?: number): boolean => {
      return (data.assignments[slot] ?? []).some((a, idx) => {
        if (slot === excludeSlot && idx === excludeIdx) return false
        return a.studentIds.includes(studentId)
      })
    },
    [data.assignments],
  )

  const handleStudentClick = useCallback(
    (slot: string, assignmentIdx: number, studentId: string) => {
      if (moving) return
      const student = data.students.find((s) => s.id === studentId)
      if (selection?.studentId === studentId && selection?.slot === slot && selection?.assignmentIdx === assignmentIdx) {
        setSelection(null)
      } else {
        setSelection({
          slot,
          assignmentIdx,
          studentId,
          studentName: student?.name ?? studentId,
        })
      }
    },
    [data.students, selection, moving],
  )

  const handleDestinationClick = useCallback(
    async (targetSlot: string, targetIdx?: number) => {
      if (!selection || moving) return
      setMoving(true)
      try {
        await onMove(selection.slot, selection.assignmentIdx, selection.studentId, targetSlot, targetIdx)
        setSelection(null)
      } finally {
        setMoving(false)
      }
    },
    [selection, moving, onMove],
  )

  // Escape key to cancel selection
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setSelection(null)
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [])

  if (weeks.length === 0) return null

  const currentWeek = weeks[Math.min(weekIdx, weeks.length - 1)]
  const fullWeek = padWeek(currentWeek)

  return (
    <div className="slot-adjust-overlay" ref={containerRef}>
      {/* Header bar */}
      <div className="slot-adjust-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <button className="btn secondary" type="button" onClick={onClose}>✕ 閉じる</button>
          <h2 style={{ margin: 0, fontSize: '1.1em' }}>コマ調整</h2>
          {selection && (
            <span className="slot-adjust-selection-badge">
              {selection.studentName} を移動中
              <button type="button" onClick={() => setSelection(null)} style={{ marginLeft: 8, background: 'none', border: 'none', cursor: 'pointer', fontWeight: 'bold' }}>✕</button>
            </span>
          )}
          {moving && <span style={{ color: '#6b7280', fontSize: '0.9em' }}>移動中...</span>}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <button
            className="btn secondary"
            type="button"
            disabled={weekIdx <= 0}
            onClick={() => setWeekIdx((i) => Math.max(0, i - 1))}
          >
            ◀ 前週
          </button>
          <span style={{ fontWeight: 600, fontSize: '0.95em' }}>
            {(() => {
              const first = currentWeek[0]
              const last = currentWeek[currentWeek.length - 1]
              const [, fm, fd] = first.split('-')
              const [, lm, ld] = last.split('-')
              return `${Number(fm)}/${Number(fd)} - ${Number(lm)}/${Number(ld)} (${weekIdx + 1}/${weeks.length})`
            })()}
          </span>
          <button
            className="btn secondary"
            type="button"
            disabled={weekIdx >= weeks.length - 1}
            onClick={() => setWeekIdx((i) => Math.min(weeks.length - 1, i + 1))}
          >
            次週 ▶
          </button>
        </div>
      </div>

      {/* Grid */}
      <div className="slot-adjust-grid">
        <table>
          <thead>
            <tr>
              <th className="sa-time-col" rowSpan={2}></th>
              {fullWeek.map((date) => {
                const [, mm, dd] = date.split('-')
                const isLecture = lectureDateSet.has(date) && !holidays.has(date)
                const dow = getIsoDayOfWeek(date)
                return (
                  <th
                    key={date}
                    colSpan={3}
                    className={`sa-day-header${!isLecture ? ' sa-day-inactive' : ''}`}
                  >
                    {Number(mm)}/{Number(dd)}({DAY_NAMES[dow]})
                  </th>
                )
              })}
            </tr>
            <tr>
              {fullWeek.map((date) => {
                const isLecture = lectureDateSet.has(date) && !holidays.has(date)
                return [
                  <th key={`${date}-t`} className={`sa-sub-header${!isLecture ? ' sa-day-inactive' : ''}`}>講師</th>,
                  <th key={`${date}-s1`} className={`sa-sub-header${!isLecture ? ' sa-day-inactive' : ''}`}>生徒</th>,
                  <th key={`${date}-s2`} className={`sa-sub-header${!isLecture ? ' sa-day-inactive' : ''}`}>生徒</th>,
                ]
              })}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: slotsPerDay }, (_, slotOff) => slotOff + 1).map((slotNumber) => (
              Array.from({ length: deskCount }, (_, deskIdx) => {
                const isFirstDesk = deskIdx === 0
                return (
                  <tr key={`${slotNumber}_${deskIdx}`} className={isFirstDesk ? 'sa-slot-first' : ''}>
                    {isFirstDesk ? (
                      <td className="sa-time-cell" rowSpan={deskCount}>{formatSlotTime(slotNumber)}</td>
                    ) : null}
                    {fullWeek.map((date) => {
                      const isLecture = lectureDateSet.has(date) && !holidays.has(date)
                      const slotKey = `${date}_${slotNumber}`
                      const assignment = isLecture ? (data.assignments[slotKey] ?? [])[deskIdx] : undefined
                      const teacher = assignment?.teacherId
                        ? instructors.find((t) => t.id === assignment.teacherId)
                        : undefined
                      const teacherName = teacher?.name ?? ''

                      // Unassigned available teacher
                      let unassignedTeacher = ''
                      if (isLecture && !assignment?.teacherId) {
                        const assignedIds = new Set(
                          (data.assignments[slotKey] ?? []).map((a) => a.teacherId).filter(Boolean),
                        )
                        // Find the deskIdx-th unassigned teacher
                        const assignedCount = (data.assignments[slotKey] ?? []).filter((a) => a.teacherId).length
                        const unassignedOffset = deskIdx - assignedCount
                        if (unassignedOffset >= 0) {
                          const unassignedList = instructors.filter(
                            (t) => !assignedIds.has(t.id) && hasInstructorAvailabilityLocal(data, isMendan, instructorPersonType, t.id, slotKey),
                          )
                          if (unassignedOffset < unassignedList.length) {
                            unassignedTeacher = unassignedList[unassignedOffset].name
                          }
                        }
                      }

                      const studentIds = assignment?.studentIds ?? []
                      const s1Id = studentIds[0]
                      const s2Id = studentIds[1]
                      const s1 = s1Id ? data.students.find((s) => s.id === s1Id) : undefined
                      const s2 = s2Id ? data.students.find((s) => s.id === s2Id) : undefined
                      const s1Subject = assignment && s1Id ? getStudentSubject(assignment, s1Id) : ''
                      const s2Subject = assignment && s2Id ? getStudentSubject(assignment, s2Id) : ''

                      // Highlight logic for selected student
                      const isSourceSlot = selection && selection.slot === slotKey
                      const isS1Source = isSourceSlot && selection.assignmentIdx === deskIdx && selection.studentId === s1Id
                      const isS2Source = isSourceSlot && selection.assignmentIdx === deskIdx && selection.studentId === s2Id

                      // Destination highlighting
                      let isAvailDest = false
                      let canAcceptHere = false
                      if (selection && isLecture) {
                        isAvailDest = isStudentAvailableForSlot(selection.studentId, slotKey)
                        // Can accept if: this slot has room and student is not already here
                        const alreadyInSlot = isStudentInSlot(selection.studentId, slotKey, selection.slot, selection.assignmentIdx)
                        if (!alreadyInSlot) {
                          if (assignment && studentIds.length < 2 && !assignment.isGroupLesson) {
                            canAcceptHere = true
                          } else if (!assignment) {
                            // Empty desk - can create new assignment if slot has room
                            const slotAssigns = data.assignments[slotKey] ?? []
                            const deskLimit = data.settings.deskCount ?? 0
                            canAcceptHere = deskLimit <= 0 || slotAssigns.length < deskLimit
                          }
                        }
                      }

                      const inactiveClass = !isLecture ? ' sa-inactive' : ''
                      const availClass = selection && isAvailDest && !isSourceSlot ? ' sa-avail' : ''

                      return [
                        // Teacher column
                        <td
                          key={`${date}-${deskIdx}-t`}
                          className={`sa-teacher${inactiveClass}${unassignedTeacher && !teacherName ? ' sa-teacher-unassigned' : ''}`}
                        >
                          {teacherName || unassignedTeacher}
                        </td>,
                        // Student 1
                        <td
                          key={`${date}-${deskIdx}-s1`}
                          className={`sa-student${inactiveClass}${availClass}${isS1Source ? ' sa-selected' : ''}${selection && canAcceptHere && studentIds.length <= 1 && !isS1Source && !isS2Source ? ' sa-dest' : ''}`}
                          onClick={() => {
                            if (!isLecture) return
                            if (s1Id && !selection) {
                              handleStudentClick(slotKey, deskIdx, s1Id)
                            } else if (selection && canAcceptHere && !isS1Source) {
                              void handleDestinationClick(slotKey, assignment ? deskIdx : undefined)
                            }
                          }}
                          title={s1 ? `${s1.name} (${s1.grade}) ${s1Subject}` : selection && canAcceptHere ? 'ここに移動' : ''}
                        >
                          {s1 ? (
                            <div className="sa-student-inner">
                              <span className="sa-student-name">{s1.name}</span>
                              <span className="sa-student-detail">{s1.grade}{s1Subject}</span>
                            </div>
                          ) : null}
                        </td>,
                        // Student 2
                        <td
                          key={`${date}-${deskIdx}-s2`}
                          className={`sa-student${inactiveClass}${availClass}${isS2Source ? ' sa-selected' : ''}${selection && canAcceptHere && studentIds.length < 2 && !isS1Source && !isS2Source ? ' sa-dest' : ''}`}
                          onClick={() => {
                            if (!isLecture) return
                            if (s2Id && !selection) {
                              handleStudentClick(slotKey, deskIdx, s2Id)
                            } else if (selection && canAcceptHere && !isS2Source) {
                              void handleDestinationClick(slotKey, assignment ? deskIdx : undefined)
                            }
                          }}
                          title={s2 ? `${s2.name} (${s2.grade}) ${s2Subject}` : selection && canAcceptHere ? 'ここに移動' : ''}
                        >
                          {s2 ? (
                            <div className="sa-student-inner">
                              <span className="sa-student-name">{s2.name}</span>
                              <span className="sa-student-detail">{s2.grade}{s2Subject}</span>
                            </div>
                          ) : null}
                        </td>,
                      ]
                    })}
                  </tr>
                )
              })
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}
