import type { Assignment, GroupLesson, RegularLesson, SessionData, Student, Teacher } from '../types'
import { findRegularLessonsForSlot, getIsoDayOfWeek, getSlotNumber, getStudentSubject } from './assignments'
import { canTeachSubject } from './subjects'
import XLSX from 'xlsx-js-style'

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
type MapEntry = { subject: string; isRegular: boolean; isGroupLesson: boolean; isMakeup: boolean; isManualRegular: boolean; isManualMakeup: boolean; makeupDate?: string; star1Color?: string; star2Color?: string }

function buildStudentAssignmentMap(
  student: Student,
  assignments: Record<string, Assignment[]>,
  regularLessons: RegularLesson[],
  groupLessons: GroupLesson[],
  dates: string[],
  teachers: Teacher[],
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
      const subInfo = a.regularSubstituteInfo?.[student.id]
      const manualMark = a.manualRegularMark?.[student.id]
      const isRegAtSlot = !!a.isRegular && findRegularLessonsForSlot(regularLessons, slotKey).some(r => r.studentIds.includes(student.id))
      const isRegular = isRegAtSlot || manualMark === 'regular'
      const isMakeup = !!makeupInfo || manualMark === 'makeup'

      // Star 1 color: green for regular, orange for makeup
      let star1Color: string | undefined
      if (isRegular) star1Color = '#bbf7d0'
      else if (isMakeup) star1Color = '#fef08a'

      // Star 2 color: pink for substitute, purple for unsupported substitute
      let star2Color: string | undefined
      let regTeacherId: string | undefined
      if (subInfo) {
        regTeacherId = subInfo.regularTeacherId
      } else {
        const regLesson = regularLessons.find(r => r.studentIds.includes(student.id))
        if (regLesson) regTeacherId = regLesson.teacherId
      }
      if (regTeacherId && a.teacherId && a.teacherId !== regTeacherId) {
        const curTeacher = teachers.find(t => t.id === a.teacherId)
        const isUnsupported = curTeacher ? !canTeachSubject(curTeacher.subjects, student.grade, subject) : false
        star2Color = isUnsupported ? '#c4b5fd' : '#fbcfe8'
      }

      map[key] = {
        subject,
        isRegular,
        isGroupLesson: !!a.isGroupLesson,
        isMakeup,
        isManualRegular: manualMark === 'regular',
        isManualMakeup: manualMark === 'makeup',
        makeupDate: makeupInfo?.date,
        star1Color,
        star2Color,
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
export function openStudentScheduleHtml(params: StudentScheduleParams): void {
  const { data, sessionId } = params
  const dates = getAllDatesInRange(data.settings)
  if (dates.length === 0) {
    alert('講習期間が未設定です')
    return
  }

  const holidaySet = new Set(data.settings.holidays)
  const students = [...data.students].sort((a, b) => a.name.localeCompare(b.name, 'ja'))
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
    const assignmentMap = buildStudentAssignmentMap(student, data.assignments, data.regularLessons, data.groupLessons ?? [], dates, data.teachers)
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
          const regularLabel = entry.isRegular ? '<br><span class="regular-tag">通常</span>' : ''
          // Cell background based on star badges
          let bgStyle = ''
          if (entry.star1Color && entry.star2Color) {
            bgStyle = ` style="background: linear-gradient(to right, ${entry.star1Color} 50%, ${entry.star2Color} 50%)"`
          } else if (entry.star1Color) {
            bgStyle = ` style="background: ${entry.star1Color}"`
          } else if (entry.star2Color) {
            bgStyle = ` style="background: ${entry.star2Color}"`
          }
          slotRows += `<td class="cell"${bgStyle}>${escapeHtml(entry.subject)}${regularLabel}</td>`
        } else if (isUnavailable) {
          slotRows += '<td class="cell unavailable"></td>'
        } else {
          slotRows += '<td class="cell"></td>'
        }
      }
      slotRows += '</tr>'
    }

    // 通常回数 table (always show, add empty row if no data)
    let regularTableHtml = '<table class="count-table"><tr><th colspan="2">通常回数(希望数)</th></tr>'
    if (regularSubjects.length > 0) {
      for (const sub of regularSubjects) {
        const actual = regularCounts[sub] ?? 0
        const expected = expectedRegularCounts[sub]
        const expectedLabel = expected != null ? `(${expected})` : ''
        regularTableHtml += `<tr><td class="count-label">${escapeHtml(sub)}</td><td class="count-val">${actual}${expectedLabel}</td></tr>`
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
      if (lectureCounts[sub]) {
        const desired = student.subjectSlots[sub]
        const desiredLabel = desired != null ? `(${desired})` : ''
        if (desired != null) lectureTotalDesired += desired
        lectureTableHtml += `<tr><td class="count-label">${escapeHtml(sub)}</td><td class="count-val">${lectureCounts[sub]}${desiredLabel}</td></tr>`
        hasLectureRows = true
      }
    }
    if (individualTotal > 0) {
      const totalDesiredLabel = lectureTotalDesired > 0 ? `(${lectureTotalDesired})` : ''
      lectureTableHtml += `<tr><td class="count-label">個別計</td><td class="count-val">${individualTotal}${totalDesiredLabel}</td></tr>`
      hasLectureRows = true
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
            <div>期間: ${escapeHtml(periodStr)}</div>
            <div class="student-name">生徒名: ${escapeHtml(student.name)} (${escapeHtml(student.grade)})</div>
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

  @media print { .toolbar { display: none; } body { padding: 0; background: #fff; } }
  @media screen { body { padding-top: 48px; } }
</style>
</head>
<body>
<div class="toolbar no-print">
  <button onclick="window.print()">🖨 印刷 / PDF保存</button>
  <button onclick="saveHtml()">💾 HTMLを保存</button>
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

function saveHtml() {
  var clone = document.documentElement.cloneNode(true);
  var toolbar = clone.querySelector('.toolbar');
  if (toolbar) toolbar.remove();
  var scripts = clone.querySelectorAll('script');
  scripts.forEach(function(s) { s.remove(); });
  var html = '<!DOCTYPE html>\\n<html lang="ja">' + clone.innerHTML + '</html>';
  var bodyTag = html.indexOf('</body>');
  var inject = '<div class="toolbar no-print">'
    + '<button onclick="window.print()">🖨 印刷 / PDF保存</button>'
    + '<button onclick="saveHtml()">💾 HTMLを保存</button>'
    + '<span>点線枠・振替欄をクリックして編集可。左上の校舎情報は全生徒に反映されます。</span>'
    + '</div>'
    + '<input id="logo-file-input" class="no-print" type="file" accept="image/*" style="position:fixed;left:-9999px;top:-9999px;width:1px;height:1px;opacity:0;pointer-events:none">'
    + '<script>\\n'
    + 'window.syncShared=' + window.syncShared.toString() + ';\n'
    + 'window.pickLogo=' + window.pickLogo.toString() + ';\n'
    + '(function(){var input=document.getElementById("logo-file-input");if(!input)return;input.addEventListener("change",function(){var file=input.files&&input.files[0];if(!file)return;var reader=new FileReader();reader.onload=function(ev){var src=ev&&ev.target?ev.target.result:"";if(!src)return;document.querySelectorAll(\'[data-shared="logo-box"]\').forEach(function(other){other.innerHTML=\'<img src="\'+src+\'" alt="logo">\';});};reader.readAsDataURL(file);});})();\n'
    + saveHtml.toString() + '\\n<\\/script>';
  if (bodyTag !== -1) html = html.slice(0, bodyTag) + inject + html.slice(bodyTag);
  var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = document.title + '.html';
  a.click();
  URL.revokeObjectURL(a.href);
}
</script>
</body>
</html>`

  const newWindow = window.open('', '_blank')
  if (!newWindow) {
    alert('ポップアップがブロックされました。ブラウザの設定を確認してください。')
    return
  }
  newWindow.document.write(html)
  newWindow.document.close()
  setupStudentScheduleWindow(newWindow, sessionId)
}

/** Export student schedules as Excel workbook (one sheet per student) */
export function exportStudentScheduleExcel(params: StudentScheduleParams): void {
  const { data } = params
  const dates = getAllDatesInRange(data.settings)
  if (dates.length === 0) {
    alert('講習期間が未設定です')
    return
  }

  const holidaySet = new Set(data.settings.holidays)
  const students = [...data.students].sort((a, b) => a.name.localeCompare(b.name, 'ja'))
  const sessionName = data.settings.name || ''
  const baseYear = new Date(data.settings.startDate).getFullYear()
  const reiwaYear = baseYear - 2018
  const slotsPerDay = data.settings.slotsPerDay || 5

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

  const wb = XLSX.utils.book_new()

  // Style definitions
  const headerStyle = { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center' as const, vertical: 'center' as const }, border: thinBorder() }
  const cellStyle = { font: { sz: 8 }, alignment: { horizontal: 'center' as const, vertical: 'center' as const }, border: thinBorder() }
  const unavailableStyle = { ...cellStyle, fill: { fgColor: { rgb: 'D1D5DB' } } }

  const groupStyle = { ...cellStyle, fill: { fgColor: { rgb: 'FEF3C7' } } }
  const sunStyle = { ...headerStyle, font: { bold: true, sz: 8, color: { rgb: 'DC2626' } } }
  const satStyle = { ...headerStyle, font: { bold: true, sz: 8, color: { rgb: '2563EB' } } }

  for (const student of students) {
    const assignmentMap = buildStudentAssignmentMap(student, data.assignments, data.regularLessons, data.groupLessons ?? [], dates, data.teachers)
    const { regularCounts, lectureCounts, groupCounts, individualTotal } = countLessons(assignmentMap)
    const expectedRegularCounts = countExpectedRegularLessons(student, data.regularLessons, dates, data.settings.holidays)
    const regularSubjects = [...new Set([...Object.keys(regularCounts), ...Object.keys(expectedRegularCounts)])].sort()
    const lectureSubjects = [...new Set([...Object.keys(lectureCounts), ...Object.keys(groupCounts)])].sort()
    const unavailableSet = buildUnavailableSet(student, dates, slotsPerDay, data.settings.holidays)
    const furikaeEntries = collectFurikaeEntries(student, data.assignments, dates)

    const startParts = data.settings.startDate.split('-')
    const endParts = data.settings.endDate.split('-')
    const periodStr = `${Number(startParts[1])}月${Number(startParts[2])}日 ～ ${Number(endParts[1])}月${Number(endParts[2])}日`

    const rows: XLSX.CellObject[][] = []
    const merges: { s: { r: number; c: number }; e: { r: number; c: number } }[] = []
    const colCount = 1 + dates.length // label col + date cols

    // Row 0: Title (merged)
    const titleRow: XLSX.CellObject[] = []
    titleRow[0] = { v: `R${reiwaYear}.${sessionName} 授業日程表`, t: 's', s: { font: { bold: true, sz: 14 }, alignment: { horizontal: 'center' as const } } }
    for (let c = 1; c < colCount; c++) titleRow[c] = { v: '', t: 's' }
    rows.push(titleRow)
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } })

    // Row 1: Period + Student info (merged)
    const infoRow: XLSX.CellObject[] = []
    infoRow[0] = { v: `期間: ${periodStr}　　生徒名: ${student.name} (${student.grade})`, t: 's', s: { font: { bold: true, sz: 11 }, alignment: { horizontal: 'center' as const } } }
    for (let c = 1; c < colCount; c++) infoRow[c] = { v: '', t: 's' }
    rows.push(infoRow)
    merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } })

    // Row 2: Month headers
    const monthRow: XLSX.CellObject[] = [{ v: '', t: 's', s: headerStyle }]
    let colIdx = 1
    for (const group of datesByMonth) {
      const [, mm] = group.month.split('-')
      monthRow[colIdx] = { v: `${Number(mm)}月`, t: 's', s: { font: { bold: true, sz: 10 }, alignment: { horizontal: 'center' as const }, fill: { fgColor: { rgb: 'E8E8E8' } }, border: thinBorder() } }
      if (group.dates.length > 1) {
        merges.push({ s: { r: 2, c: colIdx }, e: { r: 2, c: colIdx + group.dates.length - 1 } })
      }
      for (let i = 1; i < group.dates.length; i++) {
        monthRow[colIdx + i] = { v: '', t: 's', s: { border: thinBorder() } }
      }
      colIdx += group.dates.length
    }
    rows.push(monthRow)

    // Row 3: Date numbers
    const dateRow: XLSX.CellObject[] = [{ v: '', t: 's', s: headerStyle }]
    for (const date of dates) {
      const day = Number(date.split('-')[2])
      const isHol = holidaySet.has(date)
      dateRow.push({ v: day, t: 'n', s: isHol ? { ...headerStyle, fill: { fgColor: { rgb: 'E5E7EB' } } } : headerStyle })
    }
    rows.push(dateRow)

    // Row 4: Day-of-week
    const dowRow: XLSX.CellObject[] = [{ v: '', t: 's', s: headerStyle }]
    for (const date of dates) {
      const dow = getIsoDayOfWeek(date)
      const label = DAY_OF_WEEK_LABELS[dow]
      const isHol = holidaySet.has(date)
      let style: Record<string, unknown> = headerStyle
      if (dow === 0) style = sunStyle
      else if (dow === 6) style = satStyle
      if (isHol) style = { ...style, fill: { fgColor: { rgb: 'E5E7EB' } } }
      dowRow.push({ v: label, t: 's', s: style })
    }
    rows.push(dowRow)

    // Row 5: 集団授業
    const groupRow: XLSX.CellObject[] = [{ v: '集団授業', t: 's', s: headerStyle }]
    for (const date of dates) {
      const greyKey = `${date}_0`
      const isGrey = unavailableSet.has(greyKey)
      const dayOfWeek = getIsoDayOfWeek(date)
      const gls = (data.groupLessons ?? []).filter(gl => gl.dayOfWeek === dayOfWeek && gl.studentIds.includes(student.id))
      if (gls.length > 0) {
        groupRow.push({ v: gls.map(gl => gl.subject).join(', '), t: 's', s: isGrey ? { ...groupStyle, fill: { fgColor: { rgb: 'D1D5DB' } } } : groupStyle })
      } else {
        groupRow.push({ v: '', t: 's', s: isGrey ? unavailableStyle : cellStyle })
      }
    }
    rows.push(groupRow)

    // Rows 6..5+slotsPerDay: Slot rows
    for (let s = 1; s <= slotsPerDay; s++) {
      const timeLabel = SLOT_TIME_LABELS[s - 1] ?? `${s}限`
      const slotRow: XLSX.CellObject[] = [{ v: timeLabel, t: 's', s: { ...headerStyle, font: { bold: true, sz: 7 }, alignment: { horizontal: 'left' as const, vertical: 'center' as const } } }]
      for (const date of dates) {
        const key = `${date}_${s}`
        const entry = assignmentMap[key]
        const isUnavail = unavailableSet.has(key)
        if (entry && !entry.isGroupLesson) {
          const label = entry.isRegular ? `${entry.subject}(通常)` : entry.subject
          // Apply star badge fill colors
          const primaryColor = entry.star1Color ?? entry.star2Color
          const slotStyle = primaryColor
            ? { ...cellStyle, fill: { fgColor: { rgb: primaryColor.replace('#', '').toUpperCase() } } }
            : cellStyle
          slotRow.push({ v: label, t: 's', s: slotStyle })
        } else if (isUnavail) {
          slotRow.push({ v: '', t: 's', s: unavailableStyle })
        } else {
          slotRow.push({ v: '', t: 's', s: cellStyle })
        }
      }
      rows.push(slotRow)
    }

    // Blank row
    rows.push([{ v: '', t: 's' }])
    const bottomStartRow = rows.length

    // 振替授業 / 通常回数 / 講習回数 side by side
    // Columns: 0=振替from, 1=→, 2=振替to, 3=gap, 4=通常label, 5=通常val, 6=gap, 7=講習label, 8=講習val
    const furikaeCount = Math.max(5, furikaeEntries.length)
    const regularCount = Math.max(1, regularSubjects.length)
    const lectureRows: { label: string; val: string }[] = []
    let lectureTotalDesiredXls = 0
    for (const sub of lectureSubjects) {
      if (lectureCounts[sub]) {
        const desired = student.subjectSlots[sub]
        const desiredLabel = desired != null ? `(${desired})` : ''
        if (desired != null) lectureTotalDesiredXls += desired
        lectureRows.push({ label: sub, val: `${lectureCounts[sub]}${desiredLabel}` })
      }
    }
    if (individualTotal > 0) {
      const totalDesiredLabel = lectureTotalDesiredXls > 0 ? `(${lectureTotalDesiredXls})` : ''
      lectureRows.push({ label: '個別計', val: `${individualTotal}${totalDesiredLabel}` })
    }
    for (const sub of Object.keys(groupCounts).sort()) {
      lectureRows.push({ label: `${sub}(集団)`, val: `${groupCounts[sub]}` })
    }
    if (lectureRows.length === 0) lectureRows.push({ label: '', val: '' })

    const bottomRowCount = Math.max(furikaeCount + 1, regularCount + 1, lectureRows.length + 1)

    // Header row for bottom tables
    const bottomHeader: XLSX.CellObject[] = []
    bottomHeader[0] = { v: '振替授業', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    bottomHeader[1] = { v: '', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    bottomHeader[2] = { v: '', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    merges.push({ s: { r: bottomStartRow, c: 0 }, e: { r: bottomStartRow, c: 2 } })
    bottomHeader[3] = { v: '', t: 's' }
    bottomHeader[4] = { v: '通常回数(希望数)', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    bottomHeader[5] = { v: '', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    merges.push({ s: { r: bottomStartRow, c: 4 }, e: { r: bottomStartRow, c: 5 } })
    bottomHeader[6] = { v: '', t: 's' }
    bottomHeader[7] = { v: '講習回数(希望数)', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    bottomHeader[8] = { v: '', t: 's', s: { ...headerStyle, fill: { fgColor: { rgb: 'F5F5F5' } } } }
    merges.push({ s: { r: bottomStartRow, c: 7 }, e: { r: bottomStartRow, c: 8 } })
    rows.push(bottomHeader)

    // Data rows
    for (let i = 0; i < bottomRowCount - 1; i++) {
      const row: XLSX.CellObject[] = []
      // 振替
      if (i < furikaeEntries.length) {
        row[0] = { v: furikaeEntries[i].fromLabel, t: 's', s: cellStyle }
        row[1] = { v: '→', t: 's', s: { ...cellStyle, font: { bold: true, sz: 8 } } }
        row[2] = { v: furikaeEntries[i].toLabel, t: 's', s: cellStyle }
      } else if (i < furikaeCount) {
        row[0] = { v: '', t: 's', s: cellStyle }
        row[1] = { v: '→', t: 's', s: { ...cellStyle, font: { bold: true, sz: 8 } } }
        row[2] = { v: '', t: 's', s: cellStyle }
      } else {
        row[0] = { v: '', t: 's' }; row[1] = { v: '', t: 's' }; row[2] = { v: '', t: 's' }
      }
      row[3] = { v: '', t: 's' }
      // 通常回数
      if (i < regularSubjects.length) {
        const sub = regularSubjects[i]
        const actual = regularCounts[sub] ?? 0
        const expected = expectedRegularCounts[sub]
        const valStr = expected != null ? `${actual}(${expected})` : `${actual}`
        row[4] = { v: sub, t: 's', s: cellStyle }
        row[5] = { v: valStr, t: 's', s: cellStyle }
      } else if (i < regularCount) {
        row[4] = { v: '', t: 's', s: cellStyle }
        row[5] = { v: '', t: 's', s: cellStyle }
      } else {
        row[4] = { v: '', t: 's' }; row[5] = { v: '', t: 's' }
      }
      row[6] = { v: '', t: 's' }
      // 講習回数
      if (i < lectureRows.length) {
        row[7] = { v: lectureRows[i].label, t: 's', s: cellStyle }
        row[8] = { v: lectureRows[i].val || '', t: 's', s: cellStyle }
      } else {
        row[7] = { v: '', t: 's' }; row[8] = { v: '', t: 's' }
      }
      rows.push(row)
    }

    // 備考 row
    rows.push([])
    rows.push([{ v: '備考:', t: 's', s: { font: { bold: true, sz: 9 } } }, { v: student.memo || '', t: 's', s: { font: { sz: 9 } } }])

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(rows)
    ws['!merges'] = merges

    // Column widths
    const cols: { wch: number }[] = [{ wch: 12 }]
    for (let i = 0; i < dates.length; i++) cols.push({ wch: 4.5 })
    ws['!cols'] = cols

    // Sheet name (max 31 chars, no invalid chars)
    const sheetName = student.name.replace(/[\\/*?[\]:]/g, '').slice(0, 31) || `Student${students.indexOf(student) + 1}`
    XLSX.utils.book_append_sheet(wb, ws, sheetName)
  }

  const fileName = `生徒日程表_R${reiwaYear}_${sessionName}.xlsx`
  XLSX.writeFile(wb, fileName)
}

function thinBorder() {
  const side = { style: 'thin' as const, color: { rgb: '333333' } }
  return { top: side, bottom: side, left: side, right: side }
}
