import type { Assignment, GroupLesson, RegularLesson, SessionData, Student } from '../types'
import { findRegularLessonsForSlot, getIsoDayOfWeek, getSlotNumber, getStudentSubject } from './assignments'
import { constraintFor, getStudentRegularLessonStatus } from './constraints'
import { generateQrSvg } from './qrcode'
import { evaluateConstraintCards } from './slotConstraints'
import { canTeachSubject } from './subjects'

const SLOT_TIME_LABELS = [
  '13:00～14:30',
  '14:40～16:10',
  '16:20～17:50',
  '18:00～19:30',
  '19:40～21:10',
]

const DAY_OF_WEEK_LABELS = ['日', '月', '火', '水', '木', '金', '土']

type StudentScheduleParams = {
  data: SessionData
  getTeacherName: (id: string) => string
  sessionId?: string
  classroomId?: string
  baseUrl?: string
  sortedStudents?: Student[]
}

/** Get all dates in the range INCLUDING holidays (for display purposes) */
function getAllDatesInRange(settings: SessionData['settings']): string[] {
  if (!settings.startDate || !settings.endDate) return []
  const start = new Date(`${settings.startDate}T00:00:00`)
  const end = new Date(`${settings.endDate}T00:00:00`)
  const dates: string[] = []
  for (let cursor = new Date(start); cursor <= end; ) {
    const y = cursor.getFullYear()
    const m = String(cursor.getMonth() + 1).padStart(2, '0')
    const d = String(cursor.getDate()).padStart(2, '0')
    dates.push(`${y}-${m}-${d}`)
    cursor = new Date(cursor.getTime() + 24 * 60 * 60 * 1000)
  }
  return dates
}

/** Build a per-date+slot map of assignments for a given student */
type MapEntry = { subject: string; isRegular: boolean; isGroupLesson: boolean; isMakeup: boolean; isManualRegular: boolean; isManualMakeup: boolean; makeupDate?: string }

function buildStudentAssignmentMap(
  student: Student,
  assignments: Record<string, Assignment[]>,
  regularLessons: RegularLesson[],
  groupLessons: GroupLesson[],
  dates: string[],
) {
  const map: Record<string, MapEntry> = {}
  const dateSet = new Set(dates)

  for (const [slotKey, pairs] of Object.entries(assignments)) {
    const [date] = slotKey.split('_')
    if (!dateSet.has(date)) continue
    const slotNumber = getSlotNumber(slotKey)
    for (const a of pairs) {
      if (!a.studentIds.includes(student.id)) continue
      const subject = getStudentSubject(a, student.id)
      const key = `${date}_${slotNumber}`
      const makeupInfo = a.regularMakeupInfo?.[student.id]
      const substituteInfo = a.regularSubstituteInfo?.[student.id]
      const manualMark = a.manualRegularMark?.[student.id]
      const isRegAtSlot = !!a.isRegular && findRegularLessonsForSlot(regularLessons, slotKey).some(r => r.studentIds.includes(student.id))
      // Skip regular lessons marked as completed (振替済)
      if (isRegAtSlot && getStudentRegularLessonStatus(student, slotKey) === 'completed') continue
      const isRegular = isRegAtSlot || manualMark === 'regular' || !!substituteInfo
      const isMakeup = !!makeupInfo || manualMark === 'makeup'

      map[key] = {
        subject,
        isRegular,
        isGroupLesson: !!a.isGroupLesson,
        isMakeup,
        isManualRegular: manualMark === 'regular',
        isManualMakeup: manualMark === 'makeup',
        makeupDate: makeupInfo?.date,
      }
    }
  }

  for (const date of dates) {
    const dayOfWeek = getIsoDayOfWeek(date)
    for (const gl of groupLessons) {
      if (gl.dayOfWeek === dayOfWeek && gl.studentIds.includes(student.id)) {
        const key = `${date}_${gl.slotNumber}`
        if (!map[key]) {
          map[key] = { subject: gl.subject, isRegular: false, isGroupLesson: true, isMakeup: false, isManualRegular: false, isManualMakeup: false }
        }
      }
    }
  }

  return map
}

/** Collect furikae (makeup) entries for a student */
function collectFurikaeEntries(
  student: Student,
  assignments: Record<string, Assignment[]>,
  dates: string[],
): { fromLabel: string; toLabel: string }[] {
  const dateSet = new Set(dates)
  const entries: { fromLabel: string; toLabel: string }[] = []

  for (const [slotKey, pairs] of Object.entries(assignments)) {
    const [date] = slotKey.split('_')
    if (!dateSet.has(date)) continue
    const slotNumber = getSlotNumber(slotKey)
    for (const a of pairs) {
      if (!a.studentIds.includes(student.id)) continue
      const makeupInfo = a.regularMakeupInfo?.[student.id]
      if (!makeupInfo) continue
      const fromDate = makeupInfo.date
      const fromDow = DAY_OF_WEEK_LABELS[makeupInfo.dayOfWeek]
      const fromSlot = makeupInfo.slotNumber
      const fromLabel = fromDate
        ? `${Number(fromDate.split('-')[1])}/${Number(fromDate.split('-')[2])}(${fromDow})${fromSlot}限`
        : `${fromDow}${fromSlot}限`
      const toDow = DAY_OF_WEEK_LABELS[getIsoDayOfWeek(date)]
      const toLabel = `${Number(date.split('-')[1])}/${Number(date.split('-')[2])}(${toDow})${slotNumber}限`
      entries.push({ fromLabel, toLabel })
    }
  }
  return entries
}

/** Count regular, lecture (individual), and group lesson counts by subject */
function countLessons(assignmentMap: Record<string, MapEntry>) {
  const regularCounts: Record<string, number> = {}
  const lectureCounts: Record<string, number> = {}
  const groupCounts: Record<string, number> = {}
  let individualTotal = 0

  for (const entry of Object.values(assignmentMap)) {
    if (entry.isGroupLesson) {
      groupCounts[entry.subject] = (groupCounts[entry.subject] ?? 0) + 1
    } else if (entry.isRegular || entry.isMakeup || entry.isManualRegular || entry.isManualMakeup) {
      regularCounts[entry.subject] = (regularCounts[entry.subject] ?? 0) + 1
    } else {
      lectureCounts[entry.subject] = (lectureCounts[entry.subject] ?? 0) + 1
      individualTotal++
    }
  }
  return { regularCounts, lectureCounts, groupCounts, individualTotal }
}

/** Count expected regular lesson occurrences per subject for a student */
function countExpectedRegularLessons(
  student: Student,
  regularLessons: RegularLesson[],
  dates: string[],
  holidays: string[],
): Record<string, number> {
  const holidaySet = new Set(holidays)
  const expected: Record<string, number> = {}
  const studentRegulars = regularLessons.filter(r => r.studentIds.includes(student.id))
  for (const r of studentRegulars) {
    const subject = r.studentSubjects?.[student.id] ?? r.subject
    for (const date of dates) {
      if (holidaySet.has(date)) continue
      if (getIsoDayOfWeek(date) === r.dayOfWeek) {
        // Skip slots marked as completed (振替済)
        const slotKey = `${date}_${r.slotNumber}`
        if (getStudentRegularLessonStatus(student, slotKey) === 'completed') continue
        expected[subject] = (expected[subject] ?? 0) + 1
      }
    }
  }
  return expected
}

/** Build set of unavailable slot keys for a student */
function buildUnavailableSet(student: Student, dates: string[], slotsPerDay: number, holidays: string[]): Set<string> {
  const set = new Set<string>()
  const holidaySet = new Set(holidays)

  for (const date of student.unavailableDates) {
    for (let s = 0; s <= slotsPerDay; s++) {
      set.add(`${date}_${s}`)
    }
  }

  for (const slotKey of (student.unavailableSlots ?? [])) {
    set.add(slotKey)
  }

  for (const h of holidaySet) {
    for (let s = 0; s <= slotsPerDay; s++) {
      set.add(`${h}_${s}`)
    }
  }

  if (!student.submittedAt) {
    for (const date of dates) {
      for (let s = 0; s <= slotsPerDay; s++) {
        set.add(`${date}_${s}`)
      }
    }
  }

  return set
}

function escapeHtml(str: string): string {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function isHtmlElement(node: unknown): node is HTMLElement {
  return !!node && typeof node === 'object' && 'nodeType' in node && 'innerHTML' in node
}

function setupStudentScheduleWindow(targetWindow: Window, sessionId?: string): void {
  const scheduleWindow = targetWindow as Window & {
    syncShared?: (el: HTMLElement) => void
    pickLogo?: () => void
  }
  const doc = targetWindow.document
  const logoInput = doc.getElementById('logo-file-input') as HTMLInputElement | null

  scheduleWindow.syncShared = (el: HTMLElement) => {
    const key = el?.getAttribute?.('data-shared')
    if (!key) return
    const value = el.innerHTML
    doc.querySelectorAll(`[data-shared="${key}"]`).forEach((other) => {
      if (other !== el && isHtmlElement(other)) {
        other.innerHTML = value
      }
    })
    try { localStorage.setItem(`studentSchedule_${key}`, value) } catch { /* quota */ }
  }

  scheduleWindow.pickLogo = () => {
    if (!logoInput) return
    logoInput.value = ''
    logoInput.click()
  }

  if (logoInput) {
    logoInput.addEventListener('change', () => {
      const file = logoInput.files?.[0]
      if (!file) return
      const reader = new FileReader()
      reader.onload = (ev) => {
        const src = typeof ev.target?.result === 'string' ? ev.target.result : ''
        if (!src) return
        const imgHtml = `<img src="${src}" alt="logo">`
        doc.querySelectorAll('[data-shared="logo-box"]').forEach((other) => {
          if (isHtmlElement(other)) {
            other.innerHTML = imgHtml
          }
        })
        try { localStorage.setItem('studentSchedule_logo-box', imgHtml) } catch { /* quota */ }
      }
      reader.readAsDataURL(file)
    })
  }

  doc.querySelectorAll('[data-shared]').forEach((node) => {
    if (!isHtmlElement(node)) return
    const sync = () => scheduleWindow.syncShared?.(node)
    node.addEventListener('input', sync)
    node.addEventListener('keyup', sync)
    node.addEventListener('blur', sync)
  })

  // Per-student save for individual notes and furikae
  if (sessionId) {
    const saveStudentData = (page: Element) => {
      const studentId = page.getAttribute('data-student-id')
      if (!studentId) return
      const prefix = `studentSchedule_${sessionId}_${studentId}`
      const indiv = page.querySelector('.notes-individual')
      if (indiv && isHtmlElement(indiv)) {
        try { localStorage.setItem(`${prefix}_individual`, indiv.innerHTML) } catch { /* */ }
      }
      const furikaeCells = page.querySelectorAll('.furikae-cell')
      const vals: string[] = []
      furikaeCells.forEach((c) => { if (isHtmlElement(c)) vals.push(c.innerHTML) })
      try { localStorage.setItem(`${prefix}_furikae`, JSON.stringify(vals)) } catch { /* */ }
    }
    doc.querySelectorAll('.notes-individual').forEach((node) => {
      if (!isHtmlElement(node)) return
      const handler = () => {
        const page = node.closest('.page')
        if (page) saveStudentData(page)
      }
      node.addEventListener('input', handler)
      node.addEventListener('blur', handler)
    })
    doc.querySelectorAll('.furikae-cell').forEach((node) => {
      if (!isHtmlElement(node)) return
      const handler = () => {
        const page = node.closest('.page')
        if (page) saveStudentData(page)
      }
      node.addEventListener('input', handler)
      node.addEventListener('blur', handler)
    })
  }
}

/** Generate the HTML content for all student schedules */
export function openStudentScheduleHtml(params: StudentScheduleParams & { targetWindow?: Window | null }): Window | null {
  const { data, sessionId, classroomId, baseUrl, targetWindow } = params
  const dates = getAllDatesInRange(data.settings)
  if (dates.length === 0) {
    if (!targetWindow) alert('講習期間が未設定です')
    return null
  }

  const holidaySet = new Set(data.settings.holidays)
  const students = params.sortedStudents ?? data.students
  const sessionName = data.settings.name || ''
  const baseYear = new Date(data.settings.startDate).getFullYear()
  const reiwaYear = baseYear - 2018
  const slotsPerDay = data.settings.slotsPerDay || 5

  // Restore saved defaults from localStorage
  let savedLogo = ''
  let savedSchoolText = ''
  let savedSharedNotes = ''
  try {
    savedLogo = localStorage.getItem('studentSchedule_logo-box') ?? ''
    savedSchoolText = localStorage.getItem('studentSchedule_school-text') ?? ''
    savedSharedNotes = localStorage.getItem('studentSchedule_shared-notes') ?? ''
  } catch { /* localStorage unavailable */ }

  // Group dates by month
  const datesByMonth: { month: string; dates: string[] }[] = []
  let currentMonth = ''
  for (const date of dates) {
    const m = date.slice(0, 7)
    if (m !== currentMonth) {
      currentMonth = m
      datesByMonth.push({ month: m, dates: [] })
    }
    datesByMonth[datesByMonth.length - 1].dates.push(date)
  }

  const pagesHtml = students.map((student, idx) => {
    const assignmentMap = buildStudentAssignmentMap(student, data.assignments, data.regularLessons, data.groupLessons ?? [], dates)
    const { regularCounts, lectureCounts, groupCounts, individualTotal } = countLessons(assignmentMap)
    const expectedRegularCounts = countExpectedRegularLessons(student, data.regularLessons, dates, data.settings.holidays)
    const regularSubjects = [...new Set([...Object.keys(regularCounts), ...Object.keys(expectedRegularCounts)])].sort()
    const lectureSubjects = [...new Set([...Object.keys(lectureCounts), ...Object.keys(groupCounts)])].sort()

    const unavailableSet = buildUnavailableSet(student, dates, slotsPerDay, data.settings.holidays)

    // Period display
    const startParts = data.settings.startDate.split('-')
    const endParts = data.settings.endDate.split('-')
    const periodStr = `${Number(startParts[1])}月${Number(startParts[2])}日 ～ ${Number(endParts[1])}月${Number(endParts[2])}日`

    // Build table header
    let headerRow1 = ''
    let headerRow2 = ''
    let headerRow3 = ''

    for (const group of datesByMonth) {
      const [, mm] = group.month.split('-')
      headerRow1 += `<th colspan="${group.dates.length}" class="month-header">${Number(mm)}月</th>`
      for (const date of group.dates) {
        const day = Number(date.split('-')[2])
        const isHoliday = holidaySet.has(date)
        headerRow2 += `<th class="date-header${isHoliday ? ' holiday-col' : ''}">${day}</th>`
        const dow = getIsoDayOfWeek(date)
        const dowLabel = DAY_OF_WEEK_LABELS[dow]
        const isSun = dow === 0
        const isSat = dow === 6
        headerRow3 += `<th class="dow-header${isSun ? ' sun' : ''}${isSat ? ' sat' : ''}${isHoliday ? ' holiday-col' : ''}">${dowLabel}</th>`
      }
    }

    // Build constraint violation map for this student (slot key → violation messages)
    const violationMap: Record<string, string[]> = {}
    for (const date of dates) {
      for (let s = 1; s <= slotsPerDay; s++) {
        const key = `${date}_${s}`
        const entry = assignmentMap[key]
        if (!entry || entry.isGroupLesson) continue
        // Skip 振替 entries for constraint card evaluation
        if (entry.isMakeup) continue
        const msgs: string[] = []
        const slotAssigns = data.assignments[key] ?? []

        // Find the assignment containing this student
        const assignmentIdx = slotAssigns.findIndex(a => a.studentIds.includes(student.id) && !a.isGroupLesson)
        const assignment = assignmentIdx >= 0 ? slotAssigns[assignmentIdx] : undefined
        if (!assignment) continue
        // Skip 講師代行 entries
        if (assignment.regularSubstituteInfo?.[student.id]) continue

        // Evaluate constraint cards
        const reducedSlotAssignments = slotAssigns.flatMap((item, index) => {
          if (index !== assignmentIdx) return [item]
          const remainingIds = item.studentIds.filter(id => id !== student.id)
          if (remainingIds.length === 0) return []
          const nextStudentSubjects = Object.entries(item.studentSubjects ?? {}).reduce<Record<string, string>>((acc, [id, subject]) => {
            if (id !== student.id) acc[id] = subject
            return acc
          }, {})
          const nextMakeupInfo = Object.entries(item.regularMakeupInfo ?? {}).reduce<Record<string, { dayOfWeek: number; slotNumber: number; date?: string }>>((acc, [id, info]) => {
            if (id !== student.id) acc[id] = info
            return acc
          }, {})
          return [{ ...item, studentIds: remainingIds, ...(Object.keys(nextStudentSubjects).length > 0 ? { studentSubjects: nextStudentSubjects } : {}), ...(Object.keys(nextMakeupInfo).length > 0 ? { regularMakeupInfo: nextMakeupInfo } : {}) }]
        })

        const evalResult = evaluateConstraintCards(
          student, key,
          { ...data.assignments, [key]: reducedSlotAssignments },
          slotsPerDay, data.regularLessons, data.groupLessons ?? [], assignment.teacherId || undefined,
        )
        if (evalResult.blocked) {
          msgs.push(evalResult.blockReason ?? '制約カード違反')
        }

        // Incompatible teacher-student pair
        if (assignment.teacherId && constraintFor(data.constraints, assignment.teacherId, student.id) === 'incompatible') {
          const teacherObj = data.teachers.find(t => t.id === assignment.teacherId)
          msgs.push(`${teacherObj?.name ?? '講師'}との相性NG`)
        }

        // Subject mismatch
        if (assignment.teacherId) {
          const teacherObj = data.teachers.find(t => t.id === assignment.teacherId)
          const subj = getStudentSubject(assignment, student.id)
          if (teacherObj && subj && !canTeachSubject(teacherObj.subjects ?? [], student.grade, subj)) {
            msgs.push(`${teacherObj.name}は${subj}担当外`)
          }
        }

        // Incompatible student-student pair (2 students in same assignment)
        if (assignment.studentIds.length === 2) {
          const otherId = assignment.studentIds.find(id => id !== student.id)
          if (otherId && constraintFor(data.constraints, student.id, otherId) === 'incompatible') {
            const otherStudent = data.students.find(st => st.id === otherId)
            msgs.push(`${otherStudent?.name ?? '生徒'}との相性NG`)
          }
        }

        if (msgs.length > 0) violationMap[key] = msgs
      }
    }

    // Slot rows
    let slotRows = ''
    for (let s = 1; s <= slotsPerDay; s++) {
      const timeLabel = SLOT_TIME_LABELS[s - 1] ?? `${s}限`
      slotRows += `<tr><th class="slot-label">${timeLabel}</th>`
      for (const date of dates) {
        const key = `${date}_${s}`
        const entry = assignmentMap[key]
        const isUnavailable = unavailableSet.has(key)
        if (entry && !entry.isGroupLesson) {
          const regularLabel = (entry.isRegular || entry.isMakeup) ? '<br><span class="regular-tag">通常</span>' : ''
          const hasViolation = key in violationMap
          const titleAttr = hasViolation ? ` title="${escapeHtml(violationMap[key].join('\n'))}"` : ''
          slotRows += `<td class="cell${hasViolation ? ' violation' : ''}"${titleAttr} data-slot="${key}">${escapeHtml(entry.subject)}${regularLabel}</td>`
        } else if (isUnavailable) {
          slotRows += `<td class="cell unavailable" data-slot="${key}"></td>`
        } else {
          slotRows += `<td class="cell" data-slot="${key}"></td>`
        }
      }
      slotRows += '</tr>'
    }

    // 通常回数 table (always show, add empty row if no data)
    let regularTableHtml = '<table class="count-table"><tr><th colspan="2">通常回数(希望数)</th></tr>'
    if (regularSubjects.length > 0) {
      for (const sub of regularSubjects) {
        const actual = regularCounts[sub] ?? 0
        const hasExpected = expectedRegularCounts[sub] != null
        const expected = hasExpected ? (expectedRegularCounts[sub] ?? 0) : (actual > 0 ? 0 : undefined)
        const expectedLabel = expected != null ? `(${expected})` : ''
        const mismatch = expected != null && actual !== expected
        const cls = mismatch ? ' count-mismatch' : ''
        regularTableHtml += `<tr><td class="count-label${cls}">${escapeHtml(sub)}</td><td class="count-val${cls}">${actual}${expectedLabel}</td></tr>`
      }
    } else {
      regularTableHtml += '<tr><td class="count-label">&nbsp;</td><td class="count-val"></td></tr>'
    }
    regularTableHtml += '</table>'

    // 講習回数 table (individual subjects + 個別計 + group lessons)
    let lectureTableHtml = '<table class="count-table"><tr><th colspan="2">講習回数(希望数)</th></tr>'
    let hasLectureRows = false
    let lectureTotalDesired = 0
    for (const sub of lectureSubjects) {
      const actual = lectureCounts[sub] ?? 0
      const desired = student.subjectSlots[sub] ?? (actual > 0 ? 0 : undefined)
      const isGroup = !!groupCounts[sub]
      if (actual > 0 || (desired != null && desired > 0 && !isGroup)) {
        const desiredLabel = desired != null ? `(${desired})` : ''
        if (desired != null) lectureTotalDesired += desired
        const mismatch = desired != null && actual !== desired
        const cls = mismatch ? ' count-mismatch' : ''
        lectureTableHtml += `<tr><td class="count-label${cls}">${escapeHtml(sub)}</td><td class="count-val${cls}">${actual}${desiredLabel}</td></tr>`
        hasLectureRows = true
      }
    }
    {
      const totalDesiredLabel = lectureTotalDesired > 0 ? `(${lectureTotalDesired})` : ''
      const totalMismatch = (lectureTotalDesired > 0 || individualTotal > 0) && individualTotal !== lectureTotalDesired
      const totalCls = totalMismatch ? ' count-mismatch' : ''
      if (individualTotal > 0 || lectureTotalDesired > 0) {
        lectureTableHtml += `<tr><td class="count-label${totalCls}">個別計</td><td class="count-val${totalCls}">${individualTotal}${totalDesiredLabel}</td></tr>`
        hasLectureRows = true
      }
    }
    for (const sub of Object.keys(groupCounts).sort()) {
      lectureTableHtml += `<tr><td class="count-label">${escapeHtml(sub)}<br><span class="small">集団</span></td><td class="count-val">${groupCounts[sub]}</td></tr>`
      hasLectureRows = true
    }
    if (!hasLectureRows) {
      lectureTableHtml += '<tr><td class="count-label">&nbsp;</td><td class="count-val"></td></tr>'
    }
    lectureTableHtml += '</table>'

    // Restore saved individual notes and furikae from localStorage (session-scoped)
    let savedIndividual = ''
    let savedFurikaeVals: string[] = []
    if (sessionId) {
      const prefix = `studentSchedule_${sessionId}_${student.id}`
      try {
        savedIndividual = localStorage.getItem(`${prefix}_individual`) ?? ''
        const raw = localStorage.getItem(`${prefix}_furikae`)
        if (raw) savedFurikaeVals = JSON.parse(raw) as string[]
      } catch { /* */ }
    }

    // 振替授業 table - pre-fill with actual makeup data, editable
    const furikaeEntries = collectFurikaeEntries(student, data.assignments, dates)
    const furikaeRowCount = Math.max(5, furikaeEntries.length)
    let furikaeRowsHtml = ''
    for (let i = 0; i < furikaeRowCount; i++) {
      const entry = furikaeEntries[i]
      const savedFrom = savedFurikaeVals[i * 2] ?? ''
      const savedTo = savedFurikaeVals[i * 2 + 1] ?? ''
      const fromVal = savedFrom || (entry ? escapeHtml(entry.fromLabel) : '')
      const toVal = savedTo || (entry ? escapeHtml(entry.toLabel) : '')
      furikaeRowsHtml += `<tr><td class="furikae-cell" contenteditable="true">${fromVal}</td><td class="furikae-arrow">→</td><td class="furikae-cell" contenteditable="true">${toVal}</td></tr>`
    }
    const furikaeTableHtml = `<table class="furikae-table"><tr><th colspan="3">振替授業</th></tr>${furikaeRowsHtml}</table>`

    // QR code for student input URL
    let qrHtml = ''
    if (classroomId && sessionId && baseUrl) {
      const inputUrl = `${baseUrl}/#/c/${classroomId}/availability/${sessionId}/student/${student.id}`
      qrHtml = `<div class="qr-code">${generateQrSvg(inputUrl, 52)}</div>`
    }

    const pageBreak = idx < students.length - 1 ? ' page-break' : ''

    return `
      <div class="page${pageBreak}" data-student-id="${student.id}">
        <div class="header-row">
          <div class="header-left">
            <div class="logo-box" data-shared="logo-box" onclick="window.pickLogo()">${savedLogo}</div>
            <div class="school-text" contenteditable="true" data-shared="school-text" oninput="window.syncShared(this)" onkeyup="window.syncShared(this)" onblur="window.syncShared(this)">${savedSchoolText}</div>
          </div>
          <div class="title-center">
            <h1>R${reiwaYear}.${escapeHtml(sessionName)} 授業日程表</h1>
          </div>
          <div class="student-info">
            <div><span class="page-number no-print">${idx + 1}ページ　</span>期間: ${escapeHtml(periodStr)}</div>
            <div class="student-name-row">${qrHtml}<span class="student-name">生徒名: ${escapeHtml(student.name)} (${escapeHtml(student.grade)})</span></div>
          </div>
        </div>

        <table class="schedule-table">
          <thead>
            <tr><th class="corner" rowspan="3"></th>${headerRow1}</tr>
            <tr>${headerRow2}</tr>
            <tr>${headerRow3}</tr>
          </thead>
          <tbody>
            ${slotRows}
          </tbody>
        </table>

        <div class="bottom-area">
          <div class="notes-area">
            <div class="notes-col">
              <div class="notes-label">共通欄（全生徒に反映）</div>
              <div class="notes-shared" contenteditable="true" data-shared="shared-notes" oninput="window.syncShared(this)" onkeyup="window.syncShared(this)" onblur="window.syncShared(this)">${savedSharedNotes}</div>
            </div>
            <div class="notes-col">
              <div class="notes-label">個別欄</div>
              <div class="notes-individual" contenteditable="true">${savedIndividual || escapeHtml(student.memo || '')}</div>
            </div>
          </div>
          <div class="bottom-right">
            <div class="bottom-right-top">
              ${furikaeTableHtml}
              ${regularTableHtml}
              ${lectureTableHtml}
            </div>
          </div>
        </div>
      </div>`
  }).join('\n')

  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>生徒日程表 - ${escapeHtml(sessionName)}</title>
<style>
  @page {
    size: 297mm 210mm;
    margin: 8mm;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Hiragino Kaku Gothic ProN", "Meiryo", sans-serif; font-size: 10px; color: #333; -webkit-print-color-adjust: exact; print-color-adjust: exact; }

  .page {
    width: calc(297mm - 16mm);
    min-height: calc(210mm - 16mm);
    padding: 8px;
    position: relative;
  }
  .page-break { page-break-after: always; }

  @media print {
    html, body {
      width: 297mm;
      min-width: 297mm;
    }
    .page {
      width: calc(297mm - 16mm);
      min-height: calc(210mm - 16mm);
      padding: 0;
    }
    .no-print { display: none !important; }
  }

  @media screen {
    .page { border: 1px solid #ccc; margin-bottom: 16px; }
    body { background: #f0f0f0; padding: 16px; }
  }

  .title-area { margin-bottom: 6px; }
  h1 { font-size: 16px; text-align: center; margin-bottom: 0; }
  .header-row { display: flex; align-items: flex-start; margin-bottom: 6px; gap: 8px; position: relative; }
  .header-left {
    flex: 0 0 34%;
    max-width: 34%;
    display: flex;
    gap: 4px;
    align-items: flex-start;
    position: relative;
    z-index: 2;
  }
  .logo-box {
    flex: 1;
    min-height: 40px;
    border: 1px dashed #aaa;
    padding: 3px 5px;
    text-align: center;
    cursor: pointer;
    outline: none;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .logo-box img { max-width: 100%; max-height: 60px; display: block; }
  .logo-box:empty::before { content: 'クリックでロゴ画像を挿入'; color: #aaa; font-size: 8px; }
  .school-text {
    flex: 1;
    min-height: 40px;
    border: 1px dashed #aaa;
    padding: 3px 5px;
    font-size: 9px;
    line-height: 1.4;
    white-space: pre-wrap;
    outline: none;
  }
  .school-text:empty::before { content: '校舎名・TEL等'; color: #aaa; }
  .school-text:focus { border-color: #2563eb; }
  @media print {
    .logo-box { border: none !important; }
    .logo-box:empty::before { content: none; }
    .school-text { border: none !important; }
    .school-text:empty::before { content: none; }
  }
  .title-center {
    position: absolute;
    left: 34%;
    right: 190px;
    text-align: center;
    pointer-events: none;
    z-index: 1;
  }
  .student-info {
    flex: 0 0 190px;
    margin-left: auto;
    text-align: right;
    font-size: 11px;
    font-weight: bold;
    position: relative;
    z-index: 2;
  }
  .student-name { font-size: 14px; }
  .student-name-row { display: flex; align-items: center; justify-content: flex-end; gap: 6px; }
  .qr-code { flex-shrink: 0; line-height: 0; }
  .qr-code svg { display: block; }
  .page-number { font-size: 11px; font-weight: normal; color: #666; }
  .meta-row { display: flex; justify-content: space-between; font-size: 12px; }
  .period { font-weight: bold; }

  .schedule-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9px; }
  .schedule-table th, .schedule-table td { border: 1px solid #333; text-align: center; padding: 1px 2px; }
  .corner { width: 70px; min-width: 70px; background: #fff; }
  .month-header { background: #fff; font-weight: bold; font-size: 10px; }
  .date-header { background: #fff; font-size: 9px; }
  .dow-header { background: #fff; font-size: 8px; }
  .dow-header.sun { color: #dc2626; }
  .dow-header.sat { color: #2563eb; }
  .slot-label { background: #fff; font-size: 8px; white-space: nowrap; text-align: left; padding-left: 3px; }
  .cell { height: 22px; vertical-align: middle; font-size: 8px; color: #000; }
  .unavailable { background: #d1d5db; }
  .holiday-col { background: #e5e7eb; }
  .regular-tag { font-size: 7px; color: #000; }
  .group-cell { background: #fef3c7; font-size: 7px; }
  .violation { color: #dc2626; }
  .count-mismatch { color: #dc2626; }
  @media print { .violation { color: #000; } .count-mismatch { color: #000; } }

  .bottom-area { display: flex; gap: 8px; margin-top: 6px; }

  .notes-area {
    flex: 0 0 576px;
    width: 576px;
    display: flex;
    gap: 4px;
  }
  .notes-col {
    flex: 0 0 286px;
    width: 286px;
    display: flex;
    flex-direction: column;
  }
  .notes-shared, .notes-individual {
    flex: 1;
    min-height: 50px;
    max-height: 100px;
    border: 1px solid #333;
    padding: 4px 6px;
    font-size: 9px;
    line-height: 1.4;
    white-space: pre-wrap;
    outline: none;
    overflow: auto;
  }
  .notes-shared:focus, .notes-individual:focus { border-color: #2563eb; }
  .notes-label { font-size: 7px; color: #666; margin-bottom: 1px; }
  @media print { .notes-label { display: none; } }

  .bottom-right { flex: 1; display: flex; justify-content: flex-end; }
  .bottom-right-top { display: flex; gap: 6px; align-items: flex-start; }

  .count-table { border-collapse: collapse; font-size: 9px; width: 90px; }
  .count-table th { border: 1px solid #333; padding: 2px 6px; text-align: center; background: #f5f5f5; font-weight: bold; }
  .count-table td { border: 1px solid #333; padding: 1px 6px; text-align: center; }
  .count-label { text-align: left !important; white-space: nowrap; }
  .count-val { width: 28px; }
  .small { font-size: 7px; }

  .furikae-table { border-collapse: collapse; font-size: 9px; width: 170px; }
  .furikae-table th { border: 1px solid #333; padding: 2px 6px; text-align: center; background: #f5f5f5; font-weight: bold; }
  .furikae-table td { border: 1px solid #333; padding: 1px 4px; text-align: center; }
  .furikae-cell { width: 70px; height: 18px; font-size: 8px; }
  .furikae-cell:focus { outline: 1px solid #2563eb; }
  .furikae-arrow { width: 20px; font-weight: bold; }

  .toolbar {
    position: fixed; top: 0; left: 0; right: 0; background: #1e293b; color: #fff;
    padding: 8px 16px; display: flex; align-items: center; gap: 12px; z-index: 1000;
    box-shadow: 0 2px 4px rgba(0,0,0,0.3);
  }
  .toolbar button {
    background: #2563eb; color: #fff; border: none; padding: 6px 16px; border-radius: 4px;
    cursor: pointer; font-size: 13px;
  }
  .toolbar button:hover { background: #1d4ed8; }
  .toolbar span { font-size: 13px; }

  .sa-hl-source { background: #fef08a !important; outline: 2px solid #eab308 !important; outline-offset: -1px; }
  .sa-hl-student { background: #dbeafe !important; outline: 2px solid #3b82f6 !important; outline-offset: -1px; }
  @media print { .toolbar { display: none; } body { padding: 0; background: #fff; } .sa-hl-source, .sa-hl-student { background: inherit !important; outline: none !important; } }
  @media screen { body { padding-top: 48px; } }
</style>
</head>
<body>
<div class="toolbar no-print">
  <button onclick="window.print()">🖨 印刷 / PDF保存</button>
  <span>点線枠・振替欄をクリックして編集可。左上の校舎情報は全生徒に反映されます。</span>
</div>
<input id="logo-file-input" class="no-print" type="file" accept="image/*" style="position:fixed;left:-9999px;top:-9999px;width:1px;height:1px;opacity:0;pointer-events:none">
${pagesHtml}
<script>
window.syncShared = function(el) {
  var key = el && el.getAttribute && el.getAttribute('data-shared');
  if (!key) return;
  var val = el.innerHTML;
  document.querySelectorAll('[data-shared="' + key + '"]').forEach(function(other) {
    if (other !== el) other.innerHTML = val;
  });
};

window.pickLogo = function() {
  var input = document.getElementById('logo-file-input');
  if (!input) return;
  input.value = '';
  input.click();
};

(function() {
  var input = document.getElementById('logo-file-input');
  if (!input) return;
  input.addEventListener('change', function() {
    var file = input.files && input.files[0];
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function(ev) {
      var src = ev && ev.target ? ev.target.result : '';
      if (!src) return;
      document.querySelectorAll('[data-shared="logo-box"]').forEach(function(other) {
        other.innerHTML = '<img src="' + src + '" alt="logo">';
      });
    };
    reader.readAsDataURL(file);
  });
})();

</script>
</body>
</html>`

  const win = targetWindow && !targetWindow.closed ? targetWindow : window.open('', '_blank')
  if (!win) {
    alert('ポップアップがブロックされました。ブラウザの設定を確認してください。')
    return null
  }
  win.document.open()
  win.document.write(html)
  win.document.close()
  setupStudentScheduleWindow(win, sessionId)
  return win
}
