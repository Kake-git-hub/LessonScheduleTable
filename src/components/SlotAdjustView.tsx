import { useMemo, useState, useEffect, useCallback, useRef } from 'react'
import type { Assignment, SessionData, Student, Teacher } from '../types'
import { findRegularLessonsForSlot, getSlotNumber, getIsoDayOfWeek, getSlotDayOfWeek, getStudentSubject } from '../utils/assignments'
import { hasAvailability, isStudentAvailable } from '../utils/constraints'
import { teachableBaseSubjects } from '../utils/subjects'

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
    targetTeacherId?: string,
  ) => Promise<void>
  onAddStudent: (slot: string, idx: number, studentId: string, teacherId?: string, subject?: string) => Promise<void>
  onCreateStudent: (name: string, grade: string) => Promise<string>
  onUndo: () => Promise<void>
  onRedo: () => Promise<void>
  undoCount: number
  redoCount: number
  onRemoveStudent: (slot: string, assignmentIdx: number, studentId: string) => Promise<void>
  onClose: () => void
  onSelectionChange?: (sel: { slot: string; studentId: string } | null) => void
}

type SelectionInfo = {
  slot: string
  assignmentIdx: number
  studentId: string
  studentName: string
}

type StudentPickerInfo = {
  slot: string
  deskIdx: number
  /** When set, picker shows subject selection for this student */
  selectedStudentId?: string
  /** When true, picker shows new student creation form */
  creatingStudent?: boolean
}

const GRADE_OPTIONS = ['小1', '小2', '小3', '小4', '小5', '小6', '中1', '中2', '中3', '高1', '高2', '高3']

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

/** Compute star badge for a student in an assignment (same rules as main grid) */
function getStudentBadge(
  data: SessionData,
  assignment: Assignment,
  studentId: string,
  slot: string,
  isMendan: boolean,
): { star1Color: string; star1Bg: string; star2Color?: string; star2Bg?: string } | null {
  if (isMendan || !studentId) return null

  const isRegAtSlot = assignment.isRegular && findRegularLessonsForSlot(data.regularLessons, slot).some(r => r.studentIds.includes(studentId))
  const mkInfo = assignment.regularMakeupInfo?.[studentId]
  const subInfo = assignment.regularSubstituteInfo?.[studentId]
  const manualMark = assignment.manualRegularMark?.[studentId]

  // Star 1: green=通常, orange=振替
  let hasStar1 = false
  let star1Bg = '#22c55e'
  let star1Color = '#052e16'
  if (isRegAtSlot) {
    hasStar1 = true // green
  } else if (mkInfo) {
    hasStar1 = true
    star1Bg = '#eab308'
    star1Color = '#422006'
  } else if (manualMark === 'regular') {
    hasStar1 = true
  } else if (manualMark === 'makeup') {
    hasStar1 = true
    star1Bg = '#eab308'
    star1Color = '#422006'
  }

  if (!hasStar1) return null

  // Star 2: substitute teacher
  let regTeacherId: string | undefined
  if (subInfo) {
    regTeacherId = subInfo.regularTeacherId
  } else {
    const slotDow = getSlotDayOfWeek(slot)
    const slotNum = getSlotNumber(slot)
    const regLesson = data.regularLessons.find(r => r.studentIds.includes(studentId) && r.dayOfWeek === slotDow && r.slotNumber === slotNum)
    if (regLesson) regTeacherId = regLesson.teacherId
  }

  let star2Bg: string | undefined
  let star2Color: string | undefined
  if (regTeacherId && assignment.teacherId && assignment.teacherId !== regTeacherId) {
    // pink=代行, purple=担当外
    star2Bg = '#fbcfe8'
    star2Color = '#9d174d'
  }

  return { star1Color, star1Bg, star2Color, star2Bg }
}

// ---- component ----

export default function SlotAdjustView({
  data,
  instructors,
  instructorPersonType,
  isMendan,
  onMove,
  onAddStudent,
  onCreateStudent,
  onUndo,
  onRedo,
  undoCount,
  redoCount,
  onRemoveStudent,
  onClose,
  onSelectionChange,
}: SlotAdjustViewProps) {
  const [selection, setSelection] = useState<SelectionInfo | null>(null)
  const [weekIdx, setWeekIdx] = useState(0)
  const [busy, setBusy] = useState(false)
  const [studentPicker, setStudentPicker] = useState<StudentPickerInfo | null>(null)
  const [newStudentName, setNewStudentName] = useState('')
  const [newStudentGrade, setNewStudentGrade] = useState('中1')
  const pickerRef = useRef<HTMLDivElement>(null)
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
      if (busy) return
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
      setStudentPicker(null)
    },
    [data.students, selection, busy],
  )

  const runBusy = useCallback(async (fn: () => Promise<void>) => {
    setBusy(true)
    try {
      await fn()
    } finally {
      setBusy(false)
    }
  }, [])

  const handleDestinationClick = useCallback(
    async (targetSlot: string, targetIdx?: number, targetTeacherId?: string) => {
      if (!selection || busy) return
      await runBusy(async () => {
        await onMove(selection.slot, selection.assignmentIdx, selection.studentId, targetSlot, targetIdx, targetTeacherId)
        setSelection(null)
      })
    },
    [selection, busy, onMove, runBusy],
  )

  const handleAddStudent = useCallback(
    async (slot: string, deskIdx: number, studentId: string, teacherId?: string, subject?: string) => {
      setStudentPicker(null)
      await runBusy(async () => {
        await onAddStudent(slot, deskIdx, studentId, teacherId, subject)
      })
    },
    [onAddStudent, runBusy],
  )

  // Notify parent when selection changes (for schedule HTML highlighting)
  useEffect(() => {
    onSelectionChange?.(selection ? { slot: selection.slot, studentId: selection.studentId } : null)
  }, [selection, onSelectionChange])

  // Escape key to cancel selection / close picker
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        if (studentPicker) setStudentPicker(null)
        else setSelection(null)
      }
    }
    const win = containerRef.current?.ownerDocument?.defaultView ?? window
    win.addEventListener('keydown', handler)
    return () => win.removeEventListener('keydown', handler)
  }, [studentPicker])

  // Build list of students for the picker
  const pickerStudents: Student[] = useMemo(() => {
    if (!studentPicker) return []
    const slot = studentPicker.slot
    const existingIds = new Set(
      (data.assignments[slot] ?? []).flatMap(a => a.studentIds),
    )
    return data.students.filter(s => !existingIds.has(s.id))
  }, [studentPicker, data.assignments, data.students])

  // Click outside to close student picker
  useEffect(() => {
    if (!studentPicker) return
    const handler = (e: MouseEvent) => {
      if (pickerRef.current && !pickerRef.current.contains(e.target as Node)) {
        setStudentPicker(null)
      }
    }
    const doc = containerRef.current?.ownerDocument ?? document
    doc.addEventListener('mousedown', handler)
    return () => doc.removeEventListener('mousedown', handler)
  }, [studentPicker])

  if (weeks.length === 0) return null

  const currentWeek = weeks[Math.min(weekIdx, weeks.length - 1)]
  const fullWeek = padWeek(currentWeek)

  return (
    <div className="slot-adjust-overlay" ref={containerRef}>
      {/* Busy overlay */}
      {busy && <div className="sa-busy-overlay"><span className="sa-busy-spinner">処理中...</span></div>}

      {/* Header bar */}
      <div className="slot-adjust-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <button className="btn secondary" type="button" onClick={onClose}>✕ 閉じる</button>
          <h2 style={{ margin: 0, fontSize: '1.1em' }}>コマ調整</h2>
          <button className="btn secondary" type="button" disabled={undoCount === 0 || busy} onClick={() => void runBusy(onUndo)} title="元に戻す">
            ↩ 戻す
          </button>
          <button className="btn secondary" type="button" disabled={redoCount === 0 || busy} onClick={() => void runBusy(onRedo)} title="やり直し">
            ↪ やり直し
          </button>
          {selection && (() => {
            const [sDate, sSlot] = selection.slot.split('_')
            const [, sm, sd] = sDate.split('-')
            const isOtherWeek = !currentWeek.includes(sDate)
            return (
              <span className="slot-adjust-selection-badge">
                {selection.studentName} を移動中（{Number(sm)}/{Number(sd)} {sSlot}限目）
                {isOtherWeek && <span style={{ marginLeft: 6, fontSize: '0.85em', color: '#b91c1c' }}>※別の週から</span>}
                <button type="button" onClick={() => setSelection(null)} style={{ marginLeft: 8, background: 'none', border: 'none', cursor: 'pointer', fontWeight: 'bold' }}>✕</button>
              </span>
            )
          })()}
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
                      const slotAssigns = data.assignments[slotKey] ?? []
                      const assignment = isLecture ? slotAssigns[deskIdx] : undefined
                      const teacher = assignment?.teacherId
                        ? instructors.find((t) => t.id === assignment.teacherId)
                        : undefined
                      const teacherName = teacher?.name ?? ''

                      // Unassigned available teacher
                      let unassignedTeacher = ''
                      let unassignedTeacherId = ''
                      if (isLecture && !assignment?.teacherId) {
                        const assignedIds = new Set(
                          slotAssigns.map((a) => a.teacherId).filter(Boolean),
                        )
                        const assignedCount = slotAssigns.filter((a) => a.teacherId).length
                        const unassignedOffset = deskIdx - assignedCount
                        if (unassignedOffset >= 0) {
                          const unassignedList = instructors.filter(
                            (t) => !assignedIds.has(t.id) && hasInstructorAvailabilityLocal(data, isMendan, instructorPersonType, t.id, slotKey),
                          )
                          if (unassignedOffset < unassignedList.length) {
                            unassignedTeacher = unassignedList[unassignedOffset].name
                            unassignedTeacherId = unassignedList[unassignedOffset].id
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

                      // Star badges
                      const s1Badge = assignment && s1Id ? getStudentBadge(data, assignment, s1Id, slotKey, isMendan) : null
                      const s2Badge = assignment && s2Id ? getStudentBadge(data, assignment, s2Id, slotKey, isMendan) : null

                      // Highlight logic for selected student
                      const isSourceSlot = selection && selection.slot === slotKey
                      const isS1Source = isSourceSlot && selection.assignmentIdx === deskIdx && selection.studentId === s1Id
                      const isS2Source = isSourceSlot && selection.assignmentIdx === deskIdx && selection.studentId === s2Id

                      // Destination highlighting
                      let isAvailDest = false
                      let canAcceptHere = false
                      if (selection && isLecture) {
                        isAvailDest = isStudentAvailableForSlot(selection.studentId, slotKey)
                        const alreadyInSlot = isStudentInSlot(selection.studentId, slotKey, selection.slot, selection.assignmentIdx)
                        if (!alreadyInSlot) {
                          if (assignment && studentIds.length < 2 && !assignment.isGroupLesson) {
                            canAcceptHere = true
                          } else if (!assignment) {
                            // Empty desk (including unassigned-teacher ghost desks) - allow move
                            const deskLimit = data.settings.deskCount ?? 0
                            canAcceptHere = deskLimit <= 0 || slotAssigns.length < deskLimit
                          }
                        }
                      }

                      // Can open student picker: has teacher (assigned or available), has room, no selection active
                      const hasTeacher = !!assignment?.teacherId || !!unassignedTeacherId
                      const canPickStudent = isLecture && hasTeacher && studentIds.length < 2 && !selection && !assignment?.isGroupLesson
                      const isPickerOpen = studentPicker?.slot === slotKey && studentPicker?.deskIdx === deskIdx
                      const effectiveTeacherId = assignment?.teacherId || unassignedTeacherId
                      const effectiveTeacher = effectiveTeacherId ? instructors.find(t => t.id === effectiveTeacherId) : undefined
                      // Track whether the first student cell has rendered the picker already
                      let pickerRendered = false

                      const inactiveClass = !isLecture ? ' sa-inactive' : ''
                      const availClass = selection && isAvailDest && !isSourceSlot ? ' sa-avail' : ''

                      const renderStudentCell = (
                        sId: string | undefined,
                        student: Student | undefined,
                        subject: string,
                        badge: ReturnType<typeof getStudentBadge>,
                        isSource: boolean | null | 0 | undefined,
                        cellKey: string,
                        isSecondSlot: boolean,
                      ) => {
                        const isEmpty = !sId
                        const showDest = selection && canAcceptHere && !isS1Source && !isS2Source && (isSecondSlot ? studentIds.length < 2 : studentIds.length <= 1)
                        const showPicker = isEmpty && canPickStudent
                        // Only render the picker dropdown once per desk (on the first empty cell)
                        const shouldRenderPicker = isPickerOpen && showPicker && !pickerRendered
                        if (shouldRenderPicker) pickerRendered = true

                        // Subject selection sub-picker
                        const pickerSelectedStudent = studentPicker?.selectedStudentId
                          ? data.students.find(s => s.id === studentPicker.selectedStudentId)
                          : undefined
                        const pickerSubjects = pickerSelectedStudent && effectiveTeacher
                          ? teachableBaseSubjects(effectiveTeacher.subjects ?? [], pickerSelectedStudent.grade)
                          : []

                        return (
                          <td
                            key={cellKey}
                            className={`sa-student${inactiveClass}${availClass}${isSource ? ' sa-selected' : ''}${showDest ? ' sa-dest' : ''}${showPicker ? ' sa-pickable' : ''}`}
                            onClick={() => {
                              if (!isLecture || busy) return
                              if (sId && !selection) {
                                handleStudentClick(slotKey, deskIdx, sId)
                              } else if (selection && canAcceptHere && !isSource) {
                                void handleDestinationClick(slotKey, assignment ? deskIdx : undefined, !assignment ? unassignedTeacherId || undefined : undefined)
                              } else if (showPicker) {
                                setStudentPicker(isPickerOpen ? null : { slot: slotKey, deskIdx })
                              }
                            }}
                            onContextMenu={(e) => {
                              if (!sId || !student || busy || !assignment) return
                              e.preventDefault()
                              const ok = confirm(`${student.name} をこのコマから削除しますか？`)
                              if (ok) {
                                void runBusy(async () => {
                                  await onRemoveStudent(slotKey, deskIdx, sId)
                                  setSelection(null)
                                })
                              }
                            }}
                            title={student ? `${student.name} (${student.grade}) ${subject}` : selection && canAcceptHere ? 'ここに移動' : showPicker ? 'クリックで生徒を追加' : ''}
                            style={{ position: 'relative', overflow: shouldRenderPicker ? 'visible' : undefined }}
                          >
                            {student ? (
                              <div className="sa-student-inner">
                                {badge && (
                                  <span className="sa-badge-row">
                                    <span className="sa-star" style={{ background: badge.star1Bg, color: badge.star1Color }}>★</span>
                                    {badge.star2Bg && <span className="sa-star" style={{ background: badge.star2Bg, color: badge.star2Color }}>★</span>}
                                  </span>
                                )}
                                <span className="sa-student-name">{student.name}</span>
                                <span className="sa-student-detail">{student.grade}{subject}</span>
                              </div>
                            ) : showPicker ? (
                              <span className="sa-add-hint">＋</span>
                            ) : null}
                            {/* Student picker dropdown — only rendered once per desk */}
                            {shouldRenderPicker && (
                              <div className="sa-picker" ref={pickerRef} onClick={(e) => e.stopPropagation()}>
                                {studentPicker?.creatingStudent ? (
                                  <>
                                    <div className="sa-picker-title" style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                                      <button type="button" className="sa-picker-back" onClick={() => { setStudentPicker({ slot: slotKey, deskIdx }); setNewStudentName(''); setNewStudentGrade('中1') }}>◀</button>
                                      新規生徒追加
                                    </div>
                                    <div style={{ padding: '6px 8px', display: 'flex', flexDirection: 'column', gap: 4 }}>
                                      <input
                                        type="text"
                                        placeholder="生徒名"
                                        value={newStudentName}
                                        onChange={(e) => setNewStudentName(e.target.value)}
                                        style={{ fontSize: 11, padding: '3px 6px', border: '1px solid #cbd5e1', borderRadius: 4, width: '100%' }}
                                        autoFocus
                                        onKeyDown={(e) => { if (e.key === 'Enter' && newStudentName.trim()) { void (async () => { await onCreateStudent(newStudentName.trim(), newStudentGrade); setNewStudentName(''); setNewStudentGrade('中1'); setStudentPicker({ slot: slotKey, deskIdx }) })() } }}
                                      />
                                      <select
                                        value={newStudentGrade}
                                        onChange={(e) => setNewStudentGrade(e.target.value)}
                                        style={{ fontSize: 11, padding: '3px 6px', border: '1px solid #cbd5e1', borderRadius: 4 }}
                                      >
                                        {GRADE_OPTIONS.map(g => <option key={g} value={g}>{g}</option>)}
                                      </select>
                                      <button
                                        type="button"
                                        disabled={!newStudentName.trim() || busy}
                                        style={{ fontSize: 11, padding: '4px 8px', background: '#2563eb', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer' }}
                                        onClick={() => {
                                          if (!newStudentName.trim()) return
                                          void (async () => {
                                            await onCreateStudent(newStudentName.trim(), newStudentGrade)
                                            setNewStudentName('')
                                            setNewStudentGrade('中1')
                                            setStudentPicker({ slot: slotKey, deskIdx })
                                          })()
                                        }}
                                      >
                                        追加
                                      </button>
                                    </div>
                                  </>
                                ) : pickerSelectedStudent ? (
                                  <>
                                    <div className="sa-picker-title" style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                                      <button type="button" className="sa-picker-back" onClick={() => setStudentPicker({ slot: slotKey, deskIdx })}>◀</button>
                                      {pickerSelectedStudent.name} の科目
                                    </div>
                                    <div className="sa-picker-list">
                                      {pickerSubjects.map(subj => (
                                        <div
                                          key={subj}
                                          className="sa-picker-item"
                                          onClick={() => void handleAddStudent(slotKey, deskIdx, pickerSelectedStudent.id, !assignment?.teacherId ? unassignedTeacherId || undefined : undefined, subj)}
                                        >
                                          {subj}
                                        </div>
                                      ))}
                                      {pickerSubjects.length === 0 && <div className="sa-picker-empty">担当可能な科目がありません</div>}
                                    </div>
                                  </>
                                ) : (
                                  <>
                                    <div className="sa-picker-title">生徒を追加</div>
                                    <div className="sa-picker-list">
                                      <div
                                        className="sa-picker-item sa-picker-new"
                                        onClick={() => setStudentPicker({ slot: slotKey, deskIdx, creatingStudent: true })}
                                      >
                                        ＋ 新規生徒追加
                                      </div>
                                      {pickerStudents.map((st) => {
                                        const slotDow = getSlotDayOfWeek(slotKey)
                                        const slotNum = getSlotNumber(slotKey)
                                        const hasRegular = data.regularLessons.some(r => r.studentIds.includes(st.id) && r.dayOfWeek === slotDow && r.slotNumber === slotNum)
                                        const hasRegularAny = data.regularLessons.some(r => r.studentIds.includes(st.id))
                                        return (
                                          <div
                                            key={st.id}
                                            className="sa-picker-item"
                                            onClick={() => {
                                              const subjects = effectiveTeacher ? teachableBaseSubjects(effectiveTeacher.subjects ?? [], st.grade) : []
                                              if (subjects.length <= 1) {
                                                void handleAddStudent(slotKey, deskIdx, st.id, !assignment?.teacherId ? unassignedTeacherId || undefined : undefined, subjects[0])
                                              } else {
                                                setStudentPicker({ slot: slotKey, deskIdx, selectedStudentId: st.id })
                                              }
                                            }}
                                          >
                                            <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                                              {hasRegular && <span className="sa-picker-badge sa-picker-badge-slot">通常</span>}
                                              {!hasRegular && hasRegularAny && <span className="sa-picker-badge sa-picker-badge-other">通常有</span>}
                                              {st.name} <span className="sa-picker-grade">({st.grade})</span>
                                            </span>
                                          </div>
                                        )
                                      })}
                                      {pickerStudents.length === 0 && <div className="sa-picker-empty">追加可能な生徒がいません</div>}
                                    </div>
                                  </>
                                )}
                              </div>
                            )}
                          </td>
                        )
                      }

                      return [
                        // Teacher column
                        <td
                          key={`${date}-${deskIdx}-t`}
                          className={`sa-teacher${inactiveClass}${unassignedTeacher && !teacherName ? ' sa-teacher-unassigned' : ''}`}
                        >
                          {teacherName || unassignedTeacher}
                        </td>,
                        // Student 1
                        renderStudentCell(s1Id, s1, s1Subject, s1Badge, isS1Source, `${date}-${deskIdx}-s1`, false),
                        // Student 2
                        renderStudentCell(s2Id, s2, s2Subject, s2Badge, isS2Source, `${date}-${deskIdx}-s2`, true),
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
