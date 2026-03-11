import type { Assignment, GroupLesson, RegularLesson, SessionData, Student } from '../types'
import { getDatesInRange, getIsoDayOfWeek, getSlotNumber, getStudentSubject } from './assignments'

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

/** Build a per-date+slot map of assignments for a given student */
function buildStudentAssignmentMap(
  student: Student,
  assignments: Record<string, Assignment[]>,
  _regularLessons: RegularLesson[],
  groupLessons: GroupLesson[],
  dates: string[],
) {
  const map: Record<string, { subject: string; isRegular: boolean; isGroupLesson: boolean }> = {}
  const dateSet = new Set(dates)

  for (const [slotKey, pairs] of Object.entries(assignments)) {
    const [date] = slotKey.split('_')
    if (!dateSet.has(date)) continue
    const slotNumber = getSlotNumber(slotKey)
    for (const a of pairs) {
      if (!a.studentIds.includes(student.id)) continue
      const subject = getStudentSubject(a, student.id)
      const key = `${date}_${slotNumber}`
      map[key] = {
        subject,
        isRegular: !!a.isRegular,
        isGroupLesson: !!a.isGroupLesson,
      }
    }
  }

  for (const date of dates) {
    const dayOfWeek = getIsoDayOfWeek(date)
    for (const gl of groupLessons) {
      if (gl.dayOfWeek === dayOfWeek && gl.studentIds.includes(student.id)) {
        const key = `${date}_${gl.slotNumber}`
        if (!map[key]) {
          map[key] = { subject: gl.subject, isRegular: false, isGroupLesson: true }
        }
      }
    }
  }

  return map
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

  // Student-marked unavailable dates → all slots on that date
  for (const date of student.unavailableDates) {
    for (let s = 0; s <= slotsPerDay; s++) {
      set.add(`${date}_${s}`)
    }
  }

  // Student-marked unavailable individual slots
  for (const slotKey of (student.unavailableSlots ?? [])) {
    set.add(slotKey)
  }

  // Holidays that fall within the full date range (already excluded from getDatesInRange,
  // but we still want to show them grayed if the table includes surrounding dates)
  for (const h of holidaySet) {
    for (let s = 0; s <= slotsPerDay; s++) {
      set.add(`${h}_${s}`)
    }
  }

  // Also gray out all slots for dates where student is unavailable (entire date)
  // Check per-date unavailability: if student has not submitted, all slots are unavailable
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
  const dates = getDatesInRange(data.settings)
  if (dates.length === 0) {
    alert('講習期間が未設定です')
    return
  }

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
        headerRow2 += `<th class="date-header">${day}</th>`
        const dow = getIsoDayOfWeek(date)
        const dowLabel = DAY_OF_WEEK_LABELS[dow]
        const isSun = dow === 0
        const isSat = dow === 6
        headerRow3 += `<th class="dow-header${isSun ? ' sun' : ''}${isSat ? ' sat' : ''}">${dowLabel}</th>`
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

    // 振替授業 table (5 empty arrow rows)
    const furikaeRows = Array.from({ length: 5 }, () => '<tr><td class="furikae-cell"></td><td class="furikae-arrow">→</td></tr>').join('')
    const furikaeTableHtml = `<table class="furikae-table"><tr><th colspan="2">振替授業</th></tr>${furikaeRows}</table>`

    const pageBreak = idx < students.length - 1 ? ' page-break' : ''

    return `
      <div class="page${pageBreak}">
        <div class="title-area">
          <h1>R${reiwaYear}.${escapeHtml(sessionName)} 授業日程表</h1>
          <div class="meta-row">
            <span class="period">期間: ${escapeHtml(periodStr)}</span>
            <span class="student-name">生徒名: ${escapeHtml(student.name)} (${escapeHtml(student.grade)})</span>
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
  h1 { font-size: 16px; text-align: center; margin-bottom: 4px; }
  .meta-row { display: flex; justify-content: space-between; font-size: 12px; }
  .period { font-weight: bold; }
  .student-name { font-weight: bold; font-size: 14px; }

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
  .regular-tag { font-size: 7px; color: #000; }
  .group-cell { background: #fef3c7; font-size: 7px; }

  .bottom-area { display: flex; gap: 8px; margin-top: 6px; }

  .notes-box {
    flex: 0 1 auto;
    width: 180px;
    min-height: 40px;
    max-height: 80px;
    border: 2px dashed #999;
    padding: 4px;
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
  .furikae-cell { width: 50px; height: 18px; }
  .furikae-arrow { width: 24px; font-weight: bold; }

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
  <span>点線枠内をクリックして自由に編集できます。保存後もブラウザで開いて再編集可能です。</span>
</div>
${pagesHtml}
<script>
function saveHtml() {
  var clone = document.documentElement.cloneNode(true);
  var toolbar = clone.querySelector('.toolbar');
  if (toolbar) toolbar.remove();
  var scripts = clone.querySelectorAll('script');
  scripts.forEach(function(s) { s.remove(); });
  var html = '<!DOCTYPE html>\\n<html lang="ja">' + clone.innerHTML + '</html>';
  // Re-inject toolbar and save script into the saved file so it can be saved again
  var bodyTag = html.indexOf('</body>');
  var inject = '<div class="toolbar no-print">'
    + '<button onclick="window.print()">🖨 印刷 / PDF保存</button>'
    + '<button onclick="saveHtml()">💾 HTMLを保存</button>'
    + '<span>点線枠内をクリックして自由に編集できます。保存後もブラウザで開いて再編集可能です。</span>'
    + '</div>'
    + '<script>\\n' + saveHtml.toString() + '\\n<\\/script>';
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
