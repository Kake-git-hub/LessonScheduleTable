import type { Assignment, RegularLesson, SessionData, Student, Teacher } from '../types'
import { findRegularLessonsForSlot, getIsoDayOfWeek, getSlotNumber, getStudentSubject } from './assignments'
import { constraintFor, getStudentRegularLessonStatus } from './constraints'
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

type TeacherScheduleParams = {
  data: SessionData
  getStudentName: (id: string) => string
  getStudentGrade: (id: string) => string
}

/** Get all dates in the range INCLUDING holidays */
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

type SlotEntry = {
  students: { id: string; name: string; grade: string; subject: string; isRegular: boolean; isMakeup: boolean; isSubstitute: boolean }[]
  isGroupLesson: boolean
  groupSubject?: string
}

/** Build a per-date+slot map of assignments for a given teacher */
function buildTeacherAssignmentMap(
  teacher: Teacher,
  assignments: Record<string, Assignment[]>,
  regularLessons: RegularLesson[],
  dates: string[],
  allStudents: Student[],
  getStudentName: (id: string) => string,
  getStudentGrade: (id: string) => string,
): Record<string, SlotEntry> {
  const map: Record<string, SlotEntry> = {}
  const dateSet = new Set(dates)

  for (const [slotKey, pairs] of Object.entries(assignments)) {
    const [date] = slotKey.split('_')
    if (!dateSet.has(date)) continue
    const slotNumber = getSlotNumber(slotKey)
    for (const a of pairs) {
      if (a.teacherId !== teacher.id) continue
      const key = `${date}_${slotNumber}`
      if (a.isGroupLesson) {
        map[key] = { students: [], isGroupLesson: true, groupSubject: a.subject }
        continue
      }
      const students = a.studentIds
        .filter(sid => {
          // Skip students whose regular lesson at this slot is completed (振替済)
          if (a.isRegular) {
            const isRegAtSlot = findRegularLessonsForSlot(regularLessons, slotKey).some(r => r.studentIds.includes(sid))
            if (isRegAtSlot) {
              const student = allStudents.find(s => s.id === sid)
              if (student && getStudentRegularLessonStatus(student, slotKey) === 'completed') return false
            }
          }
          return true
        })
        .map(sid => {
          const subject = getStudentSubject(a, sid)
          const isRegular = !!a.isRegular
          const isMakeup = !!a.regularMakeupInfo?.[sid]
          const isSubstitute = !!a.regularSubstituteInfo?.[sid]
          return {
            id: sid,
            name: getStudentName(sid),
            grade: getStudentGrade(sid),
            subject,
            isRegular,
            isMakeup,
            isSubstitute,
          }
        })
      if (students.length > 0) {
        map[key] = { students, isGroupLesson: false }
      }
    }
  }

  return map
}

/** Build set of unavailable slot keys for a teacher */
function buildTeacherUnavailableSet(
  teacher: Teacher,
  availability: SessionData['availability'],
  dates: string[],
  slotsPerDay: number,
  holidays: string[],
): Set<string> {
  const set = new Set<string>()
  const holidaySet = new Set(holidays)
  const availKey = `teacher_${teacher.id}`
  const availableSlots = new Set(availability[availKey] ?? [])

  for (const date of dates) {
    const isHoliday = holidaySet.has(date)
    for (let s = 1; s <= slotsPerDay; s++) {
      const slotKey = `${date}_${s}`
      if (isHoliday || !availableSlots.has(slotKey)) {
        set.add(slotKey)
      }
    }
  }

  return set
}

function escapeHtml(str: string): string {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

/** Count assigned slots for a teacher */
function countTeacherSlots(assignmentMap: Record<string, SlotEntry>): { regular: number; lecture: number; group: number; total: number } {
  let regular = 0
  let lecture = 0
  let group = 0
  for (const entry of Object.values(assignmentMap)) {
    if (entry.isGroupLesson) {
      group++
    } else {
      // Check if any student in the slot is regular, makeup, or substitute
      const hasRegular = entry.students.some(s => s.isRegular || s.isMakeup || s.isSubstitute)
      if (hasRegular) regular++
      else lecture++
    }
  }
  return { regular, lecture, group, total: regular + lecture + group }
}

/** Generate the HTML content for all teacher schedules */
export function openTeacherScheduleHtml(params: TeacherScheduleParams & { targetWindow?: Window | null }): Window | null {
  const { data, getStudentName, getStudentGrade, targetWindow } = params
  const dates = getAllDatesInRange(data.settings)
  if (dates.length === 0) {
    if (!targetWindow) alert('講習期間が未設定です')
    return null
  }

  const holidaySet = new Set(data.settings.holidays)
  const teachers = [...data.teachers].sort((a, b) => a.name.localeCompare(b.name, 'ja'))
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

  const pagesHtml = teachers.map((teacher, idx) => {
    const assignmentMap = buildTeacherAssignmentMap(teacher, data.assignments, data.regularLessons, dates, data.students, getStudentName, getStudentGrade)
    const unavailableSet = buildTeacherUnavailableSet(teacher, data.availability, dates, slotsPerDay, data.settings.holidays)
    const counts = countTeacherSlots(assignmentMap)

    // Period display
    const startParts = data.settings.startDate.split('-')
    const endParts = data.settings.endDate.split('-')
    const periodStr = `${Number(startParts[1])}月${Number(startParts[2])}日 ～ ${Number(endParts[1])}月${Number(endParts[2])}日`

    // Regular lesson info for this teacher
    const regularLessons = data.regularLessons.filter(r => r.teacherId === teacher.id)
    const regularInfo = regularLessons.map(r => {
      const studentNames = r.studentIds.map(sid => getStudentName(sid)).join('・')
      const subj = r.studentSubjects ? Object.values(r.studentSubjects).join('/') : r.subject
      return `${DAY_OF_WEEK_LABELS[r.dayOfWeek]}${r.slotNumber}限 ${studentNames}(${subj})`
    }).join('、')

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

    // Build constraint violation map for this teacher (slot key → violation messages)
    const violationMap: Record<string, string[]> = {}
    for (const date of dates) {
      for (let s = 1; s <= slotsPerDay; s++) {
        const key = `${date}_${s}`
        const entry = assignmentMap[key]
        if (!entry || entry.isGroupLesson) continue
        const msgs: string[] = []
        const slotAssignments = data.assignments[key] ?? []
        const assignmentIdx = slotAssignments.findIndex(a => a.teacherId === teacher.id && !a.isGroupLesson)
        const assignment = assignmentIdx >= 0 ? slotAssignments[assignmentIdx] : undefined
        if (!assignment) continue

        for (const st of entry.students) {
          // Skip 振替 and 講師代行 students for constraint card evaluation
          if (st.isMakeup || st.isSubstitute) continue
          const student = data.students.find(x => x.id === st.id)
          if (!student) continue

          // Evaluate constraint cards (same approach as getManualConstraintWarnings)
          const reducedSlotAssignments = slotAssignments.flatMap((item, index) => {
            if (index !== assignmentIdx) return [item]
            const remainingIds = item.studentIds.filter(id => id !== st.id)
            if (remainingIds.length === 0) return []
            const nextStudentSubjects = Object.entries(item.studentSubjects ?? {}).reduce<Record<string, string>>((acc, [id, subject]) => {
              if (id !== st.id) acc[id] = subject
              return acc
            }, {})
            const nextMakeupInfo = Object.entries(item.regularMakeupInfo ?? {}).reduce<Record<string, { dayOfWeek: number; slotNumber: number; date?: string }>>((acc, [id, info]) => {
              if (id !== st.id) acc[id] = info
              return acc
            }, {})
            return [{ ...item, studentIds: remainingIds, ...(Object.keys(nextStudentSubjects).length > 0 ? { studentSubjects: nextStudentSubjects } : {}), ...(Object.keys(nextMakeupInfo).length > 0 ? { regularMakeupInfo: nextMakeupInfo } : {}) }]
          })

          const evalResult = evaluateConstraintCards(
            student, key,
            { ...data.assignments, [key]: reducedSlotAssignments },
            slotsPerDay, data.regularLessons, data.groupLessons ?? [], assignment.teacherId,
          )
          if (evalResult.blocked) {
            msgs.push(`${st.name}: ${evalResult.blockReason ?? '制約カード違反'}`)
          }

          // Incompatible teacher-student pair
          if (assignment.teacherId && constraintFor(data.constraints, assignment.teacherId, st.id) === 'incompatible') {
            msgs.push(`${st.name}: 講師との相性NG`)
          }

          // Subject mismatch
          const teacherObj = data.teachers.find(t => t.id === assignment.teacherId)
          if (teacherObj && st.subject) {
            if (!canTeachSubject(teacherObj.subjects ?? [], student.grade, st.subject)) {
              msgs.push(`${st.name}: ${teacherObj.name}は${st.subject}担当外`)
            }
          }
        }

        // Incompatible student-student pair (2 students)
        if (entry.students.length === 2) {
          const [s1, s2] = entry.students
          if (constraintFor(data.constraints, s1.id, s2.id) === 'incompatible') {
            msgs.push(`${s1.name}と${s2.name}: 生徒間相性NG`)
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
        const isHoliday = holidaySet.has(date)
        const hasViolation = key in violationMap
        if (entry) {
          const studentIds = entry.students.map(st => st.id).join(',')
          if (entry.isGroupLesson) {
            slotRows += `<td class="cell group-cell" data-slot="${key}">\u25A0${escapeHtml(entry.groupSubject ?? '')}</td>`
          } else {
            const lines = entry.students.map(st => {
              const tag = st.isRegular ? '<span class="regular-tag">\u901A\u5E38</span>' : st.isMakeup ? '<span class="makeup-tag">\u632F\u66FF</span>' : ''
              const subTag = st.isSubstitute ? '<span class="substitute-tag">\u4EE3\u884C</span>' : ''
              return `<span class="student-line">${escapeHtml(st.name)}<br>${escapeHtml(st.grade)}${escapeHtml(st.subject)}${tag}${subTag}</span>`
            }).join('<hr class="cell-divider">')
            const titleAttr = hasViolation ? ` title="${escapeHtml(violationMap[key].join('\n'))}"` : ''
            slotRows += `<td class="cell assigned-cell${hasViolation ? ' violation' : ''}"${titleAttr} data-slot="${key}" data-student-ids="${studentIds}">${lines}</td>`
          }
        } else if (isUnavailable || isHoliday) {
          slotRows += `<td class="cell unavailable" data-slot="${key}"></td>`
        } else {
          slotRows += `<td class="cell" data-slot="${key}"></td>`
        }
      }
      slotRows += '</tr>'
    }

    // Counts summary
    const countsHtml = `<table class="count-table"><tr><th colspan="2">コマ数</th></tr>
      <tr><td class="count-label">通常</td><td class="count-val">${counts.regular}</td></tr>
      <tr><td class="count-label">講習</td><td class="count-val">${counts.lecture}</td></tr>
      ${counts.group > 0 ? `<tr><td class="count-label">集団</td><td class="count-val">${counts.group}</td></tr>` : ''}
      <tr><td class="count-label"><b>合計</b></td><td class="count-val"><b>${counts.total}</b></td></tr>
    </table>`

    const pageBreak = idx < teachers.length - 1 ? ' page-break' : ''

    return `
      <div class="page${pageBreak}">
        <div class="header-row">
          <div class="title-center-teacher">
            <h1>R${reiwaYear}.${escapeHtml(sessionName)} 講師日程表</h1>
          </div>
          <div class="teacher-info">
            <div><span class="page-number no-print">${idx + 1}ページ　</span>期間: ${escapeHtml(periodStr)}</div>
            <div class="teacher-name">講師名: ${escapeHtml(teacher.name)}</div>
            <div class="teacher-subjects">担当科目: ${escapeHtml(teacher.subjects.join(', '))}</div>
          </div>
        </div>
        ${regularInfo ? `<div class="regular-info">通常授業: ${escapeHtml(regularInfo)}</div>` : ''}

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
          <div class="teacher-notes-area">
            <div class="notes-label">備考</div>
            <div class="notes-box" contenteditable="true">${escapeHtml(teacher.memo || '')}</div>
          </div>
          <div class="bottom-right">
            ${countsHtml}
          </div>
        </div>
      </div>`
  }).join('\n')

  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<title>講師日程表 - ${escapeHtml(sessionName)}</title>
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

  h1 { font-size: 16px; text-align: center; margin-bottom: 0; }
  .header-row { display: flex; align-items: flex-start; margin-bottom: 4px; gap: 8px; position: relative; }
  .title-center-teacher {
    flex: 1;
    text-align: center;
  }
  .teacher-info {
    flex: 0 0 220px;
    text-align: right;
    font-size: 11px;
    font-weight: bold;
  }
  .teacher-name { font-size: 14px; }
  .page-number { font-size: 11px; font-weight: normal; color: #666; margin-left: 8px; }
  .teacher-subjects { font-size: 10px; font-weight: normal; color: #555; }
  .regular-info { font-size: 9px; color: #555; margin-bottom: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

  .schedule-table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 9px; }
  .schedule-table th, .schedule-table td { border: 1px solid #333; text-align: center; padding: 1px 2px; }
  .corner { width: 70px; min-width: 70px; background: #fff; }
  .month-header { background: #fff; font-weight: bold; font-size: 10px; }
  .date-header { background: #fff; font-size: 9px; }
  .dow-header { background: #fff; font-size: 8px; }
  .dow-header.sun { color: #dc2626; }
  .dow-header.sat { color: #2563eb; }
  .slot-label { background: #fff; font-size: 8px; white-space: nowrap; text-align: left; padding-left: 3px; }
  .cell { height: 36px; vertical-align: middle; font-size: 7px; color: #000; line-height: 1.2; overflow: hidden; }
  .assigned-cell { font-size: 7px; padding: 1px; text-align: left; }
  .unavailable { background: #d1d5db; }
  .holiday-col { background: #e5e7eb; }
  .regular-tag { font-size: 6px; color: #16a34a; font-weight: bold; margin-left: 1px; }
  .makeup-tag { font-size: 6px; color: #d97706; font-weight: bold; margin-left: 1px; }
  .substitute-tag { font-size: 6px; color: #dc2626; font-weight: bold; margin-left: 1px; }
  .group-cell { background: #fef3c7; font-size: 7px; font-weight: bold; }
  .student-line { display: block; }
  .cell-divider { border: none; border-top: 1px dashed #999; margin: 1px 0; }
  .violation { color: #dc2626; }
  @media print { .violation { color: #000; } }

  .bottom-area { display: flex; gap: 8px; margin-top: 6px; }
  .teacher-notes-area { flex: 1; display: flex; flex-direction: column; }
  .notes-box {
    flex: 1;
    min-height: 40px;
    max-height: 60px;
    border: 1px solid #333;
    padding: 4px 6px;
    font-size: 9px;
    line-height: 1.4;
    white-space: pre-wrap;
    outline: none;
    overflow: auto;
  }
  .notes-box:focus { border-color: #2563eb; }
  .notes-label { font-size: 7px; color: #666; margin-bottom: 1px; }
  @media print { .notes-label { display: none; } }

  .bottom-right { flex: 0 0 auto; display: flex; justify-content: flex-end; }

  .count-table { border-collapse: collapse; font-size: 9px; width: 90px; }
  .count-table th { border: 1px solid #333; padding: 2px 6px; text-align: center; background: #f5f5f5; font-weight: bold; }
  .count-table td { border: 1px solid #333; padding: 1px 6px; text-align: center; }
  .count-label { text-align: left !important; white-space: nowrap; }
  .count-val { width: 28px; }

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

  @media print { .toolbar { display: none; } body { padding: 0; background: #fff; } .sa-hl-source, .sa-hl-student { background: inherit !important; outline: none !important; } }
  @media screen { body { padding-top: 48px; } }

  /* Slot adjust highlighting */
  .sa-hl-source { background: #fef08a !important; outline: 2px solid #eab308 !important; outline-offset: -1px; }
  .sa-hl-student { background: #dbeafe !important; outline: 2px solid #3b82f6 !important; outline-offset: -1px; }
</style>
</head>
<body>
<div class="toolbar no-print">
  <button onclick="window.print()">🖨 印刷 / PDF保存</button>
  <span>講師日程表 — 備考欄をクリックして編集可</span>
</div>
${pagesHtml}
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
  return win
}
