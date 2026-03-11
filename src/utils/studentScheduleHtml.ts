import type { Assignment, GroupLesson, RegularLesson, SessionData, Student } from '../types'
import { getIsoDayOfWeek, getSlotNumber, getStudentSubject } from './assignments'

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
function buildStudentAssignmentMap(
  student: Student,
  assignments: Record<string, Assignment[]>,
  _regularLessons: RegularLesson[],
  groupLessons: GroupLesson[],
  dates: string[],
) {
  const map: Record<string, { subject: string; isRegular: boolean; isGroupLesson: boolean; isMakeup: boolean; makeupDate?: string }> = {}
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
      map[key] = {
        subject,
        isRegular: !!a.isRegular,
        isGroupLesson: !!a.isGroupLesson,
        isMakeup: !!makeupInfo,
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
          map[key] = { subject: gl.subject, isRegular: false, isGroupLesson: true, isMakeup: false }
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
function countLessons(assignmentMap: Record<string, { subject: string; isRegular: boolean; isGroupLesson: boolean }>) {
  const regularCounts: Record<string, number> = {}
  const lectureCounts: Record<string, number> = {}
  const groupCounts: Record<string, number> = {}
  let individualTotal = 0

  for (const entry of Object.values(assignmentMap)) {
    if (entry.isGroupLesson) {
      groupCounts[entry.subject] = (groupCounts[entry.subject] ?? 0) + 1
    } else if (entry.isRegular) {
      regularCounts[entry.subject] = (regularCounts[entry.subject] ?? 0) + 1
    } else {
      lectureCounts[entry.subject] = (lectureCounts[entry.subject] ?? 0) + 1
      individualTotal++
    }
  }
  return { regularCounts, lectureCounts, groupCounts, individualTotal }
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

/** Generate the HTML content for all student schedules */
export function openStudentScheduleHtml(params: StudentScheduleParams): void {
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

  const pagesHtml = students.map((student, idx) => {
    const assignmentMap = buildStudentAssignmentMap(student, data.assignments, data.regularLessons, data.groupLessons ?? [], dates)
    const { regularCounts, lectureCounts, groupCounts, individualTotal } = countLessons(assignmentMap)
    const regularSubjects = [...new Set(Object.keys(regularCounts))].sort()
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

    // 集団授業 row
    let groupRow = ''
    for (const date of dates) {
      const greyKey = `${date}_0`
      const isGrey = unavailableSet.has(greyKey)
      const dayOfWeek = getIsoDayOfWeek(date)
      const groupLessonsOnDate = (data.groupLessons ?? []).filter(
        (gl) => gl.dayOfWeek === dayOfWeek && gl.studentIds.includes(student.id),
      )
      if (groupLessonsOnDate.length > 0) {
        groupRow += `<td class="cell group-cell${isGrey ? ' unavailable' : ''}">${escapeHtml(groupLessonsOnDate.map((gl) => gl.subject).join(', '))}</td>`
      } else {
        groupRow += `<td class="cell${isGrey ? ' unavailable' : ''}"></td>`
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
          slotRows += `<td class="cell">${escapeHtml(entry.subject)}${regularLabel}</td>`
        } else if (isUnavailable) {
          slotRows += '<td class="cell unavailable"></td>'
        } else {
          slotRows += '<td class="cell"></td>'
        }
      }
      slotRows += '</tr>'
    }

    // 通常回数 table
    let regularTableHtml = ''
    if (regularSubjects.length > 0) {
      regularTableHtml = '<table class="count-table"><tr><th colspan="2">通常回数</th></tr>'
      for (const sub of regularSubjects) {
        regularTableHtml += `<tr><td class="count-label">${escapeHtml(sub)}</td><td class="count-val">${regularCounts[sub] ?? 0}</td></tr>`
      }
      regularTableHtml += '</table>'
    }

    // 講習回数 table (individual subjects + 個別計 + group lessons)
    let lectureTableHtml = '<table class="count-table"><tr><th colspan="2">講習回数</th></tr>'
    for (const sub of lectureSubjects) {
      if (lectureCounts[sub]) {
        lectureTableHtml += `<tr><td class="count-label">${escapeHtml(sub)}</td><td class="count-val">${lectureCounts[sub]}</td></tr>`
      }
    }
    if (individualTotal > 0) {
      lectureTableHtml += `<tr><td class="count-label">個別計</td><td class="count-val">${individualTotal}</td></tr>`
    }
    for (const sub of Object.keys(groupCounts).sort()) {
      lectureTableHtml += `<tr><td class="count-label">${escapeHtml(sub)}<br><span class="small">集団</span></td><td class="count-val">${groupCounts[sub]}</td></tr>`
    }
    lectureTableHtml += '</table>'

    // 振替授業 table - pre-fill with actual makeup data, editable
    const furikaeEntries = collectFurikaeEntries(student, data.assignments, dates)
    const furikaeRowCount = Math.max(5, furikaeEntries.length)
    let furikaeRowsHtml = ''
    for (let i = 0; i < furikaeRowCount; i++) {
      const entry = furikaeEntries[i]
      const fromVal = entry ? escapeHtml(entry.fromLabel) : ''
      const toVal = entry ? escapeHtml(entry.toLabel) : ''
      furikaeRowsHtml += `<tr><td class="furikae-cell" contenteditable="true">${fromVal}</td><td class="furikae-arrow">→</td><td class="furikae-cell" contenteditable="true">${toVal}</td></tr>`
    }
    const furikaeTableHtml = `<table class="furikae-table"><tr><th colspan="3">振替授業</th></tr>${furikaeRowsHtml}</table>`

    const pageBreak = idx < students.length - 1 ? ' page-break' : ''

    return `
      <div class="page${pageBreak}">
        <div class="header-row">
          <div class="school-info" contenteditable="true" data-shared="school-info"></div>
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
            <tr><th class="slot-label">集団授業</th>${groupRow}</tr>
            ${slotRows}
          </tbody>
        </table>

        <div class="bottom-area">
          <div class="notes-box" contenteditable="true">${escapeHtml(student.memo || '')}</div>
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
    size: A4 landscape;
    margin: 8mm;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Hiragino Kaku Gothic ProN", "Meiryo", sans-serif; font-size: 10px; color: #333; }

  .page { padding: 8px; position: relative; }
  .page-break { page-break-after: always; }

  @media print {
    .page { padding: 0; }
    .no-print { display: none !important; }
  }

  @media screen {
    .page { border: 1px solid #ccc; margin-bottom: 16px; max-width: 297mm; min-height: 200mm; }
    body { background: #f0f0f0; padding: 16px; }
  }

  .title-area { margin-bottom: 6px; }
  h1 { font-size: 16px; text-align: center; margin-bottom: 0; }
  .header-row { display: flex; align-items: flex-start; margin-bottom: 6px; gap: 8px; }
  .school-info {
    flex: 0 0 160px;
    min-height: 36px;
    border: 1px dashed #aaa;
    padding: 3px 5px;
    font-size: 9px;
    line-height: 1.4;
    white-space: pre-wrap;
    outline: none;
  }
  .school-info:empty::before {
    content: 'ロゴ・校舎名・TEL等';
    color: #aaa;
  }
  .school-info:focus { border-color: #2563eb; }
  .title-center { flex: 1; text-align: center; }
  .student-info { flex: 0 0 auto; text-align: right; font-size: 11px; font-weight: bold; }
  .student-name { font-size: 14px; }
  .meta-row { display: flex; justify-content: space-between; font-size: 12px; }
  .period { font-weight: bold; }

  .schedule-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9px; }
  .schedule-table th, .schedule-table td { border: 1px solid #333; text-align: center; padding: 1px 2px; }
  .corner { width: 70px; min-width: 70px; background: #f5f5f5; }
  .month-header { background: #e8e8e8; font-weight: bold; font-size: 10px; }
  .date-header { background: #f5f5f5; font-size: 9px; }
  .dow-header { background: #f5f5f5; font-size: 8px; }
  .dow-header.sun { color: #dc2626; }
  .dow-header.sat { color: #2563eb; }
  .slot-label { background: #f5f5f5; font-size: 8px; white-space: nowrap; text-align: left; padding-left: 3px; }
  .cell { height: 22px; vertical-align: middle; font-size: 8px; color: #000; }
  .unavailable { background: #d1d5db; }
  .holiday-col { background: #e5e7eb; }
  .regular-tag { font-size: 7px; color: #000; }
  .group-cell { background: #fef3c7; font-size: 7px; }

  .bottom-area { display: flex; gap: 8px; margin-top: 6px; }

  .notes-box {
    flex: 1 1 50%;
    min-height: 50px;
    max-height: 100px;
    border: 2px dashed #999;
    padding: 4px 6px;
    font-size: 9px;
    line-height: 1.4;
    white-space: pre-wrap;
    outline: none;
    overflow: auto;
  }
  .notes-box:focus { border-color: #2563eb; }

  .bottom-right { flex: 1; display: flex; justify-content: flex-end; }
  .bottom-right-top { display: flex; gap: 6px; align-items: flex-start; }

  .count-table { border-collapse: collapse; font-size: 9px; }
  .count-table th { border: 1px solid #333; padding: 2px 6px; text-align: center; background: #f5f5f5; font-weight: bold; }
  .count-table td { border: 1px solid #333; padding: 1px 6px; text-align: center; }
  .count-label { text-align: left !important; white-space: nowrap; }
  .count-val { width: 28px; }
  .small { font-size: 7px; }

  .furikae-table { border-collapse: collapse; font-size: 9px; }
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
${pagesHtml}
<script>
// Sync school-info across all pages
(function() {
  var infos = document.querySelectorAll('[data-shared="school-info"]');
  infos.forEach(function(el) {
    el.addEventListener('input', function() {
      var val = el.innerHTML;
      infos.forEach(function(other) { if (other !== el) other.innerHTML = val; });
    });
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
    + '<script>\\n'
    + '(function(){var infos=document.querySelectorAll(\\'[data-shared="school-info"]\\');infos.forEach(function(el){el.addEventListener("input",function(){var val=el.innerHTML;infos.forEach(function(o){if(o!==el)o.innerHTML=val;});});});})();\\n'
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
}
