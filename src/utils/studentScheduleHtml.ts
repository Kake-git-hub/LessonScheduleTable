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
  // Map: "YYYY-MM-DD_slotNumber" → { subject, isRegular, isGroupLesson }
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

  // Also check group lessons auto-placed on each date
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

/** Count regular and lecture lesson counts by subject */
function countLessons(assignmentMap: Record<string, { subject: string; isRegular: boolean; isGroupLesson: boolean }>) {
  const regularCounts: Record<string, number> = {}
  const lectureCounts: Record<string, number> = {}

  for (const entry of Object.values(assignmentMap)) {
    if (entry.isGroupLesson) continue
    if (entry.isRegular) {
      regularCounts[entry.subject] = (regularCounts[entry.subject] ?? 0) + 1
    } else {
      lectureCounts[entry.subject] = (lectureCounts[entry.subject] ?? 0) + 1
    }
  }
  return { regularCounts, lectureCounts }
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

  // Group dates by month
  const datesByMonth: { month: string; dates: string[] }[] = []
  let currentMonth = ''
  for (const date of dates) {
    const m = date.slice(0, 7) // YYYY-MM
    if (m !== currentMonth) {
      currentMonth = m
      datesByMonth.push({ month: m, dates: [] })
    }
    datesByMonth[datesByMonth.length - 1].dates.push(date)
  }

  const slotsPerDay = data.settings.slotsPerDay || 5

  // Build HTML
  const pagesHtml = students.map((student, idx) => {
    const assignmentMap = buildStudentAssignmentMap(student, data.assignments, data.regularLessons, data.groupLessons ?? [], dates)
    const { regularCounts, lectureCounts } = countLessons(assignmentMap)
    const allSubjects = [...new Set([...Object.keys(regularCounts), ...Object.keys(lectureCounts)])].sort()

    // Period display
    const startParts = data.settings.startDate.split('-')
    const endParts = data.settings.endDate.split('-')
    const periodStr = `${Number(startParts[1])}月${Number(startParts[2])}日 ～ ${Number(endParts[1])}月${Number(endParts[2])}日`

    // Build table header
    let headerRow1 = '' // month row
    let headerRow2 = '' // date row
    let headerRow3 = '' // day-of-week row

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

    // Build body rows
    // Row 0: 集団授業
    let groupRow = ''
    for (const date of dates) {
      const dayOfWeek = getIsoDayOfWeek(date)
      const groupLessonsOnDate = (data.groupLessons ?? []).filter(
        (gl) => gl.dayOfWeek === dayOfWeek && gl.studentIds.includes(student.id),
      )
      if (groupLessonsOnDate.length > 0) {
        groupRow += `<td class="cell group-cell">${escapeHtml(groupLessonsOnDate.map((gl) => gl.subject).join(', '))}</td>`
      } else {
        groupRow += '<td class="cell"></td>'
      }
    }

    // Slot rows (1-based to slotsPerDay)
    let slotRows = ''
    for (let s = 1; s <= slotsPerDay; s++) {
      const timeLabel = SLOT_TIME_LABELS[s - 1] ?? `${s}限`
      slotRows += `<tr><th class="slot-label">${timeLabel}</th>`
      for (const date of dates) {
        const key = `${date}_${s}`
        const entry = assignmentMap[key]
        if (entry && !entry.isGroupLesson) {
          const regularLabel = entry.isRegular ? '<br><span class="regular-tag">通常</span>' : ''
          const bgClass = entry.isRegular ? ' regular-bg' : ''
          slotRows += `<td class="cell${bgClass}">${escapeHtml(entry.subject)}${regularLabel}</td>`
        } else {
          slotRows += '<td class="cell"></td>'
        }
      }
      slotRows += '</tr>'
    }

    // Lesson count table
    let countTableHtml = ''
    if (allSubjects.length > 0) {
      countTableHtml = '<table class="count-table"><tr><th></th>'
      for (const sub of allSubjects) countTableHtml += `<th>${escapeHtml(sub)}</th>`
      countTableHtml += '</tr><tr><td>通常回数</td>'
      for (const sub of allSubjects) countTableHtml += `<td>${regularCounts[sub] ?? 0}</td>`
      countTableHtml += '</tr><tr><td>講習回数</td>'
      for (const sub of allSubjects) countTableHtml += `<td>${lectureCounts[sub] ?? 0}</td>`
      countTableHtml += '</tr></table>'
    }

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
          <div class="counts-area">
            ${countTableHtml}
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
  .cell { height: 22px; vertical-align: middle; font-size: 8px; }
  .regular-bg { background: #dcfce7; }
  .regular-tag { font-size: 7px; color: #16a34a; }
  .group-cell { background: #fef3c7; font-size: 7px; }

  .bottom-area { display: flex; gap: 12px; margin-top: 6px; }

  .notes-box {
    flex: 1;
    min-height: 60px;
    border: 2px dashed #999;
    padding: 6px;
    font-size: 10px;
    line-height: 1.5;
    white-space: pre-wrap;
    outline: none;
  }
  .notes-box:focus { border-color: #2563eb; }

  .counts-area { flex-shrink: 0; }
  .count-table { border-collapse: collapse; font-size: 9px; }
  .count-table th, .count-table td { border: 1px solid #333; padding: 2px 6px; text-align: center; }
  .count-table th { background: #f5f5f5; }

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
  <span>点線枠内をクリックして自由に編集できます。編集後に印刷してください。</span>
</div>
${pagesHtml}
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
