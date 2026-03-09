import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import html2canvas from 'html2canvas'
import type { Assignment } from '../types'

// ---------- Japanese font loading ----------

let cachedFontBase64: string | null = null

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer)
  const len = bytes.length
  // Build base64 without spread operator to avoid stack overflow on large files
  let binary = ''
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i])
  }
  return btoa(binary)
}

async function loadJapaneseFont(doc: jsPDF): Promise<void> {
  if (!cachedFontBase64) {
    // Try bundled font first, then CDN fallbacks (must be static TrueType TTF, NOT variable fonts)
    const base = import.meta.env.BASE_URL ?? '/'
    const urls = [
      `${base}fonts/SawarabiGothic-Regular.ttf`,
      'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/sawarabigothic/SawarabiGothic-Regular.ttf',
      'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/mplusrounded1c/MPLUSRounded1c-Regular.ttf',
    ]
    let buf: ArrayBuffer | null = null
    for (const url of urls) {
      try {
        const res = await fetch(url)
        if (res.ok) {
          buf = await res.arrayBuffer()
          if (buf.byteLength > 10000) break // sanity check: font should be >10KB
          buf = null
        }
      } catch {
        // try next
      }
    }
    if (!buf) {
      throw new Error('日本語フォントの読み込みに失敗しました。ネットワーク接続を確認してください。')
    }
    cachedFontBase64 = arrayBufferToBase64(buf)
  }
  doc.addFileToVFS('NotoSansJP-Regular.ttf', cachedFontBase64)
  doc.addFont('NotoSansJP-Regular.ttf', 'NotoSansJP', 'normal')
  doc.addFont('NotoSansJP-Regular.ttf', 'NotoSansJP', 'bold')
  doc.setFont('NotoSansJP')
}

// ---------- Schedule PDF (A3 portrait, one week per page) ----------

export type SchedulePdfParams = {
  classroomName?: string
  sessionName: string
  startDate: string
  endDate: string
  slotsPerDay: number
  deskCount?: number
  holidays: string[]
  assignments: Record<string, Assignment[]>
  baselineAssignments?: Record<string, Assignment[]>
  getTeacherName: (id: string) => string
  getStudentName: (id: string) => string
  getStudentGrade: (id: string) => string
  getStudentSubject: (a: Assignment, studentId: string) => string
  getIsoDayOfWeek: (date: string) => number
}

type StudentPdfCell = {
  text: string
  compareKey: string
  fillColor?: [number, number, number]
  textColor?: [number, number, number]
}

type DeskPdfCell = {
  teacher: string
  teacherCompareKey: string
  student1: StudentPdfCell
  student2: StudentPdfCell
}

export async function exportSchedulePdf(params: SchedulePdfParams): Promise<void> {
  const {
    classroomName, sessionName, startDate, endDate, slotsPerDay, deskCount, holidays,
    assignments, baselineAssignments, getTeacherName, getStudentName, getStudentGrade, getStudentSubject, getIsoDayOfWeek,
  } = params

  if (!startDate || !endDate) { alert('開始日・終了日を設定してください。'); return }

  const holidaySet = new Set(holidays)
  const dayNames = ['日', '月', '火', '水', '木', '金', '土']
  const effectiveDeskCount = deskCount && deskCount > 0
    ? deskCount
    : Math.max(1, ...Object.values(assignments).map((slotAssignments) => Math.max(1, slotAssignments.length)))
  const startHour = 13
  const startMinute = 0
  const slotIntervalMinutes = 100
  const pdfTextBlack = [0, 0, 0] as [number, number, number]
  const colorNormal = { fill: [22, 163, 74] as [number, number, number], text: pdfTextBlack }
  const colorMakeup = { fill: [234, 179, 8] as [number, number, number], text: pdfTextBlack }
  const colorSubstitute = { fill: [251, 207, 232] as [number, number, number], text: pdfTextBlack }
  const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value))

  const formatSlotTimeLabel = (slotNumber: number): string => {
    const totalMinutes = startHour * 60 + startMinute + (slotNumber - 1) * slotIntervalMinutes
    const hour = Math.floor(totalMinutes / 60)
    const minute = totalMinutes % 60
    return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`
  }

  const buildStudentPdfCell = (assignment: Assignment | undefined, studentId: string | undefined): StudentPdfCell => {
    if (!assignment || !studentId) return { text: '', compareKey: '' }

    const studentName = getStudentName(studentId)
    const studentGrade = getStudentGrade(studentId)
    const studentSubject = getStudentSubject(assignment, studentId)
    const text = `${studentName}\n${studentGrade}${studentSubject}`

    if (assignment.regularSubstituteInfo?.[studentId]) {
      const info = assignment.regularSubstituteInfo[studentId]
      return {
        text,
        compareKey: `${studentId}|${studentSubject}|substitute|${info.regularTeacherId}|${info.dayOfWeek}|${info.slotNumber}|${info.date ?? ''}`,
        fillColor: colorSubstitute.fill,
        textColor: colorSubstitute.text,
      }
    }
    if (assignment.regularMakeupInfo?.[studentId]) {
      const info = assignment.regularMakeupInfo[studentId]
      return {
        text,
        compareKey: `${studentId}|${studentSubject}|makeup|${info.dayOfWeek}|${info.slotNumber}|${info.date ?? ''}`,
        fillColor: colorMakeup.fill,
        textColor: colorMakeup.text,
      }
    }
    if (assignment.isRegular) {
      return {
        text,
        compareKey: `${studentId}|${studentSubject}|regular`,
        fillColor: colorNormal.fill,
        textColor: colorNormal.text,
      }
    }
    return { text, compareKey: `${studentId}|${studentSubject}|normal` }
  }

  const buildDeskPdfCell = (
    date: string,
    slotNumber: number,
    deskIndex: number,
    lectureDateSet: Set<string>,
    sourceAssignments: Record<string, Assignment[]>,
  ): DeskPdfCell => {
    if (!lectureDateSet.has(date) || holidaySet.has(date)) {
      return {
        teacher: '',
        teacherCompareKey: '',
        student1: { text: '', compareKey: '' },
        student2: { text: '', compareKey: '' },
      }
    }

    const assignment = sourceAssignments[`${date}_${slotNumber}`]?.[deskIndex]
    return {
      teacher: assignment?.teacherId ? getTeacherName(assignment.teacherId) : '',
      teacherCompareKey: assignment?.teacherId ?? '',
      student1: buildStudentPdfCell(assignment, assignment?.studentIds[0]),
      student2: buildStudentPdfCell(assignment, assignment?.studentIds[1]),
    }
  }

  const isChangedPdfCell = (
    daySubCol: number,
    currentDeskCell: DeskPdfCell,
    baselineDeskCell: DeskPdfCell | null,
  ): boolean => {
    if (!baselineDeskCell) return false
    return (
      (daySubCol === 1 && currentDeskCell.teacherCompareKey !== baselineDeskCell.teacherCompareKey)
      || (daySubCol === 2 && currentDeskCell.student1.compareKey !== baselineDeskCell.student1.compareKey)
      || (daySubCol === 3 && currentDeskCell.student2.compareKey !== baselineDeskCell.student2.compareKey)
    )
  }

  const allDates: string[] = []
  const start = new Date(`${startDate}T00:00:00`)
  const end = new Date(`${endDate}T00:00:00`)
  for (let cursor = new Date(start); cursor <= end; cursor = new Date(cursor.getTime() + 86400000)) {
    const y = cursor.getFullYear()
    const m = String(cursor.getMonth() + 1).padStart(2, '0')
    const d = String(cursor.getDate()).padStart(2, '0')
    allDates.push(`${y}-${m}-${d}`)
  }
  if (allDates.length === 0) { alert('期間に日付がありません。'); return }

  const weeks: string[][] = []
  let currentWeek: string[] = []
  for (const date of allDates) {
    const dow = getIsoDayOfWeek(date)
    if (dow === 1 && currentWeek.length > 0) {
      weeks.push(currentWeek)
      currentWeek = []
    }
    currentWeek.push(date)
  }
  if (currentWeek.length > 0) weeks.push(currentWeek)

  let doc: jsPDF
  try {
    doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a3' })
    await loadJapaneseFont(doc)
  } catch (err) {
    alert('PDF生成に失敗しました: ' + String(err))
    return
  }

  const mmPerPt = 25.4 / 72
  const pdfLineHeightFactor = 1.15
  const fitFontSizeToWidth = (
    lines: string[],
    usableWidthMm: number,
    maxFontSize: number,
    minFontSize: number,
    fontStyle: 'normal' | 'bold' = 'normal',
  ): number => {
    const filteredLines = lines.filter((line) => line.trim().length > 0)
    if (filteredLines.length === 0 || usableWidthMm <= 0) return minFontSize

    let low = minFontSize
    let high = maxFontSize
    let best = minFontSize

    while (high - low > 0.05) {
      const mid = (low + high) / 2
      doc.setFont('NotoSansJP', fontStyle)
      doc.setFontSize(mid)
      const widestLine = filteredLines.reduce((max, line) => Math.max(max, doc.getTextWidth(line)), 0)
      if (widestLine <= usableWidthMm) {
        best = mid
        low = mid
      } else {
        high = mid
      }
    }

    return clamp(best, minFontSize, maxFontSize)
  }

  const resolveHeaderHeights = (bodyRowHeight: number): { headRow1Height: number; headRow2Height: number } => ({
    headRow1Height: clamp(Math.min(5.2, bodyRowHeight * 0.9), 3.7, 5.2),
    headRow2Height: clamp(Math.min(4.8, bodyRowHeight * 0.84), 3.5, 4.8),
  })

  const fitBodyRowHeightToPage = (totalBodyRowUnits: number, availableTableHeight: number): number => {
    let low = 2.6
    let high = 8.5
    let best = low
    const safeHeight = Math.max(40, availableTableHeight - 2.2)

    while (high - low > 0.01) {
      const mid = (low + high) / 2
      const { headRow1Height, headRow2Height } = resolveHeaderHeights(mid)
      const totalHeight = totalBodyRowUnits * mid + headRow1Height + headRow2Height
      if (totalHeight <= safeHeight) {
        best = mid
        low = mid
      } else {
        high = mid
      }
    }

    return clamp(best, 2.6, 8.5)
  }

  const dowOrder = [1, 2, 3, 4, 5, 6, 0]

  for (let wi = 0; wi < weeks.length; wi++) {
    if (wi > 0) doc.addPage('a3', 'portrait')

    const weekDates = weeks[wi]
    const firstDate = weekDates[0]
    const fullWeek: string[] = []
    const firstDow = getIsoDayOfWeek(firstDate)
    const startPad = firstDow === 0 ? 6 : firstDow - 1
    const firstMs = new Date(`${firstDate}T00:00:00`).getTime()
    for (let p = startPad; p > 0; p--) {
      const d = new Date(firstMs - p * 86400000)
      fullWeek.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`)
    }
    for (const date of weekDates) fullWeek.push(date)
    const lastDateMs = new Date(`${weekDates[weekDates.length - 1]}T00:00:00`).getTime()
    let tailIdx = 1
    while (fullWeek.length < 7) {
      const d = new Date(lastDateMs + tailIdx * 86400000)
      fullWeek.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`)
      tailIdx++
    }
    const lectureDateSet = new Set(weekDates)

    const pageWidth = doc.internal.pageSize.getWidth()
    const pageHeight = doc.internal.pageSize.getHeight()
    const margin = 6
    const titlePrefix = classroomName?.trim() ? `${classroomName.trim()} ` : ''
    const [, fm, fd] = firstDate.split('-')
    const lastDate = weekDates[weekDates.length - 1]
    const [, lm, ld] = lastDate.split('-')
    const titleText = `${titlePrefix}${sessionName}  ${Number(fm)}/${Number(fd)} - ${Number(lm)}/${Number(ld)}`
    const titleTop = 7
    const titleGap = 2.5
    const titleFontSize = fitFontSizeToWidth([titleText], pageWidth - margin * 2, 14, 8, 'bold')
    const titleHeight = titleFontSize * mmPerPt * pdfLineHeightFactor
    const titleBaselineY = titleTop + titleFontSize * mmPerPt
    const tableStartY = titleTop + titleHeight + titleGap
    const tableBottomMargin = 6

    doc.setFont('NotoSansJP', 'bold')
    doc.setFontSize(titleFontSize)
    doc.text(titleText, margin, titleBaselineY)

    const headerRow1: string[] = ['']
    const headerRow2: string[] = ['']
    for (let i = 0; i < 7; i++) {
      const date = fullWeek[i]
      const [, mm, dd] = date.split('-')
      headerRow1.push(`${Number(mm)}/${Number(dd)}(${dayNames[dowOrder[i]]})`, '', '', '')
      headerRow2.push('机番', '担当講師', '生徒', '生徒')
    }

    const bodyRows: string[][] = []
    const rowSlotNum: number[] = []
    const rowDeskIdx: number[] = []
    const rowHasAnyContent: boolean[] = []
    const teacherLinesForFit: string[] = []
    const studentLinesForFit: string[] = []
    for (let slotNumber = 1; slotNumber <= slotsPerDay; slotNumber++) {
      for (let deskIdx = 0; deskIdx < effectiveDeskCount; deskIdx++) {
        const row: string[] = [deskIdx === 0 ? formatSlotTimeLabel(slotNumber) : '']
        let hasAnyContent = false
        for (let dayIdx = 0; dayIdx < 7; dayIdx++) {
          const date = fullWeek[dayIdx]
          const deskCell = buildDeskPdfCell(date, slotNumber, deskIdx, lectureDateSet, assignments)
          if (deskCell.teacher || deskCell.student1.text || deskCell.student2.text) hasAnyContent = true
          if (deskCell.teacher) teacherLinesForFit.push(deskCell.teacher)
          studentLinesForFit.push(...deskCell.student1.text.split('\n').filter(Boolean))
          studentLinesForFit.push(...deskCell.student2.text.split('\n').filter(Boolean))
          row.push(String(deskIdx + 1), deskCell.teacher, deskCell.student1.text, deskCell.student2.text)
        }
        bodyRows.push(row)
        rowSlotNum.push(slotNumber)
        rowDeskIdx.push(deskIdx)
        rowHasAnyContent.push(hasAnyContent)
      }
    }

    const availWidth = pageWidth - margin * 2
    const timeColWidth = 10
    const dayBlockWidth = (availWidth - timeColWidth) / 7
    const deskColWidth = 4.5
    const teacherColWidth = 10
    const studentColWidth = (dayBlockWidth - deskColWidth - teacherColWidth) / 2
    const columnStyles: Record<number, { cellWidth: number; halign?: 'center' | 'left' }> = {
      0: { cellWidth: timeColWidth, halign: 'center' },
    }
    for (let dayIdx = 0; dayIdx < 7; dayIdx++) {
      const baseCol = 1 + dayIdx * 4
      columnStyles[baseCol] = { cellWidth: deskColWidth, halign: 'center' }
      columnStyles[baseCol + 1] = { cellWidth: teacherColWidth, halign: 'left' }
      columnStyles[baseCol + 2] = { cellWidth: studentColWidth, halign: 'center' }
      columnStyles[baseCol + 3] = { cellWidth: studentColWidth, halign: 'center' }
    }

    const emptyRowFactor = 0.58
    const totalBodyRows = Math.max(1, rowHasAnyContent.length)
    const availableTableHeight = Math.max(120, pageHeight - tableStartY - tableBottomMargin)
    const targetBodyRowHeight = fitBodyRowHeightToPage(totalBodyRows, availableTableHeight)
    const { headRow1Height, headRow2Height } = resolveHeaderHeights(targetBodyRowHeight)
    const rowMinHeights = rowHasAnyContent.map((hasContent) => hasContent ? targetBodyRowHeight : targetBodyRowHeight * emptyRowFactor)
    const cellPadding = clamp(targetBodyRowHeight * 0.045, 0.08, 0.28)
    const emptyRowPadding = clamp(cellPadding * 0.45, 0.04, 0.12)
    const headerCellPadding = clamp(cellPadding * 0.9, 0.08, 0.22)
    const bodyUsableHeight = Math.max(1.5, targetBodyRowHeight - cellPadding * 2)
    const teacherMaxByHeight = (bodyUsableHeight / (1 * pdfLineHeightFactor)) / mmPerPt
    const studentMaxByHeight = (bodyUsableHeight / (2 * pdfLineHeightFactor)) / mmPerPt
    const teacherUsableWidth = Math.max(2, teacherColWidth - cellPadding * 2)
    const studentUsableWidth = Math.max(2, studentColWidth - cellPadding * 2)
    const teacherFontSize = fitFontSizeToWidth(
      teacherLinesForFit,
      teacherUsableWidth,
      clamp(teacherMaxByHeight, 3, 8),
      2.4,
      'normal',
    )
    const studentFontSize = fitFontSizeToWidth(
      studentLinesForFit,
      studentUsableWidth,
      clamp(studentMaxByHeight, 2.8, 7),
      2.1,
      'bold',
    )
    const timeFontSize = clamp((bodyUsableHeight / pdfLineHeightFactor) / mmPerPt, 2.8, 6)
    const baseFontSize = Math.min(teacherFontSize, studentFontSize)
    const headFontSize = clamp(Math.max(baseFontSize, 3.1), 3.1, 4.8)
    const subHeadFontSize = clamp(Math.max(baseFontSize - 0.05, 2.9), 2.9, 4.4)
    const diffBorderInset = clamp(targetBodyRowHeight * 0.06, 0.22, 0.45)
    const diffBorderWidth = clamp(targetBodyRowHeight * 0.055, 0.24, 0.45)

    const timeGroupBounds = new Map<number, { top: number; height: number }>()

    autoTable(doc, {
      startY: tableStartY,
      margin: { left: margin, right: margin },
      theme: 'grid',
      styles: {
        font: 'NotoSansJP',
        fontSize: baseFontSize,
        cellPadding,
        lineWidth: 0.2,
        lineColor: [80, 80, 80],
        valign: 'middle',
        overflow: 'linebreak',
        minCellHeight: targetBodyRowHeight,
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [40, 40, 40],
        minCellHeight: headRow1Height,
        cellPadding: headerCellPadding,
      },
      head: [headerRow1, headerRow2],
      body: bodyRows,
      columnStyles,
      didParseCell: (hookData) => {
        if (hookData.section === 'head' && hookData.row.index === 0) {
          hookData.cell.styles.fontSize = headFontSize
          hookData.cell.styles.minCellHeight = headRow1Height
          const col = hookData.column.index
          if (col > 0 && (col - 1) % 4 === 0) {
            hookData.cell.colSpan = 4
            hookData.cell.styles.halign = 'center'
            hookData.cell.styles.fontStyle = 'bold'
            const dayIdx = Math.floor((col - 1) / 4)
            const date = fullWeek[dayIdx]
            const isSunday = getIsoDayOfWeek(date) === 0
            const isHolidayColumn = holidaySet.has(date) || !lectureDateSet.has(date)
            if (isHolidayColumn) {
              hookData.cell.styles.fillColor = [229, 231, 235]
            }
            if (isSunday || holidaySet.has(date)) {
              hookData.cell.styles.textColor = [220, 38, 38]
            }
          }
        }

        if (hookData.section === 'head' && hookData.row.index === 1) {
          hookData.cell.styles.fontSize = subHeadFontSize
          hookData.cell.styles.minCellHeight = headRow2Height
          const col = hookData.column.index
          if (col > 0) {
            const dayIdx = Math.floor((col - 1) / 4)
            const date = fullWeek[dayIdx]
            const isHolidayColumn = holidaySet.has(date) || !lectureDateSet.has(date)
            hookData.cell.styles.fillColor = isHolidayColumn ? [229, 231, 235] : [245, 245, 245]
          } else {
            hookData.cell.styles.fillColor = [255, 255, 255]
          }
        }

        if (hookData.section === 'body') {
          const slotNumber = rowSlotNum[hookData.row.index] ?? 1
          const deskIdx = rowDeskIdx[hookData.row.index] ?? 0
          const col = hookData.column.index
          const rowHasContent = rowHasAnyContent[hookData.row.index] ?? true
          const rowMinHeight = rowMinHeights[hookData.row.index] ?? targetBodyRowHeight
          hookData.cell.styles.minCellHeight = rowMinHeight
          hookData.cell.styles.cellPadding = rowHasContent ? cellPadding : emptyRowPadding

          if (col === 0) {
            hookData.cell.text = ['']
            hookData.cell.styles.fontStyle = 'bold'
            hookData.cell.styles.halign = 'center'
            hookData.cell.styles.valign = 'middle'
            hookData.cell.styles.fillColor = [255, 255, 255]
            hookData.cell.styles.textColor = [0, 0, 0]
            hookData.cell.styles.lineWidth = {
              top: deskIdx === 0 ? 0.2 : 0,
              right: 0.2,
              bottom: deskIdx === effectiveDeskCount - 1 ? 0.2 : 0,
              left: 0.2,
            }
            hookData.cell.styles.fontSize = timeFontSize
            return
          }

          const dayIdx = Math.floor((col - 1) / 4)
          const dayDate = fullWeek[dayIdx]
          const daySubCol = (col - 1) % 4
          const deskCell = buildDeskPdfCell(dayDate, slotNumber, deskIdx, lectureDateSet, assignments)
          const isHolidayColumn = holidaySet.has(dayDate) || !lectureDateSet.has(dayDate)

          if (isHolidayColumn) {
            hookData.cell.styles.fillColor = [229, 231, 235]
          }

          if (daySubCol === 0) {
            hookData.cell.styles.halign = 'center'
            hookData.cell.styles.fontStyle = 'bold'
            if (!lectureDateSet.has(dayDate) || holidaySet.has(dayDate)) {
              hookData.cell.styles.textColor = [148, 163, 184]
            }
          }

          if (daySubCol === 1) {
            hookData.cell.styles.fontSize = teacherFontSize
            hookData.cell.styles.halign = 'center'
            hookData.cell.styles.valign = 'middle'
          }

          if (daySubCol === 2 || daySubCol === 3) {
            const studentCell = daySubCol === 2 ? deskCell.student1 : deskCell.student2
            hookData.cell.styles.fontSize = studentFontSize
            hookData.cell.styles.halign = 'center'
            if (studentCell.fillColor) {
              hookData.cell.styles.fillColor = studentCell.fillColor
              hookData.cell.styles.textColor = studentCell.textColor ?? [0, 0, 0]
              hookData.cell.styles.fontStyle = 'bold'
            }
          }
        }
      },
      didDrawCell: (hookData) => {
        if (hookData.section !== 'body') return

        const col = hookData.column.index
        const slotNumber = rowSlotNum[hookData.row.index] ?? 1
        const deskIdx = rowDeskIdx[hookData.row.index] ?? 0

        if (col === 0) {
          const existingBounds = timeGroupBounds.get(slotNumber)
          if (existingBounds) {
            existingBounds.height += hookData.cell.height
          } else {
            timeGroupBounds.set(slotNumber, { top: hookData.cell.y, height: hookData.cell.height })
          }

          if (deskIdx !== effectiveDeskCount - 1) return

          const bounds = timeGroupBounds.get(slotNumber)
          if (!bounds) return

          const centerX = hookData.cell.x + hookData.cell.width / 2
          const centerY = bounds.top + bounds.height / 2
          doc.setFont('NotoSansJP', 'bold')
          doc.setFontSize(timeFontSize)
          doc.setTextColor(0, 0, 0)
          doc.text(formatSlotTimeLabel(slotNumber), centerX, centerY, { align: 'center', baseline: 'middle' })
          return
        }

        const dayIdx = Math.floor((col - 1) / 4)
        const dayDate = fullWeek[dayIdx]
        const daySubCol = (col - 1) % 4

        if (daySubCol === 0) return

        const deskCell = buildDeskPdfCell(dayDate, slotNumber, deskIdx, lectureDateSet, assignments)
        const baselineDeskCell = baselineAssignments
          ? buildDeskPdfCell(dayDate, slotNumber, deskIdx, lectureDateSet, baselineAssignments)
          : null

        if (!isChangedPdfCell(daySubCol, deskCell, baselineDeskCell)) return

        const inset = diffBorderInset
        const rectX = hookData.cell.x + inset
        const rectY = hookData.cell.y + inset
        const rectW = Math.max(0, hookData.cell.width - inset * 2)
        const rectH = Math.max(0, hookData.cell.height - inset * 2)

        doc.setDrawColor(220, 38, 38)
        doc.setLineWidth(diffBorderWidth)
        doc.rect(rectX, rectY, rectW, rectH)
      },
    })
  }

  try {
    const now = new Date()
    const ts = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`
    doc.save(`コマ割り_${sessionName}_${ts}.pdf`)
  } catch (err) {
    console.error('PDF save error:', err)
    alert('PDF出力に失敗しました: ' + String(err))
  }
}

// ---------- Email receipt PDF ----------

export type EmailReceiptPdfParams = {
  sessionName: string
  recipientName: string
  emailType: string
  sentAt: string
}

export async function downloadEmailReceiptPdf(params: EmailReceiptPdfParams): Promise<void> {
  const { sessionName, recipientName, emailType, sentAt } = params

  let doc: jsPDF
  try {
    doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })
    await loadJapaneseFont(doc)
  } catch (err) {
    alert('PDF生成に失敗しました: ' + String(err))
    return
  }

  const pageWidth = doc.internal.pageSize.getWidth()
  const centerX = pageWidth / 2

  // Title
  doc.setFontSize(18)
  doc.text('メール送信記録', centerX, 30, { align: 'center' })

  // Content
  doc.setFontSize(12)
  const lines = [
    `以下の内容でメールを送信しました。`,
    '',
    `■ セッション名: ${sessionName}`,
    `■ 送信先: ${recipientName}`,
    `■ 送信内容: ${emailType}`,
    `■ 送信日時: ${sentAt}`,
    '',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '',
    '【重要】この書類は送信記録として大切に保管してください。',
    '',
    '　この PDF は、上記の内容でメールを送信した旨を記録するものです。',
    '　紛失しないよう、適切なフォルダに保存してください。',
    '　後日確認が必要になった場合に備え、削除しないことを推奨します。',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
  ]

  let y = 50
  for (const line of lines) {
    doc.text(line, 20, y)
    y += 8
  }

  // Footer
  doc.setFontSize(9)
  doc.setTextColor(120)
  doc.text(`発行日: ${new Date().toLocaleDateString('ja-JP')}`, centerX, 270, { align: 'center' })

  doc.save(`メール送信記録_${recipientName}_${emailType}.pdf`)
}

// ---------- Submission receipt PDF ----------

export type SubmissionReceiptPdfParams = {
  sessionName: string
  personName: string
  personType: '講師' | '生徒' | '保護者'
  submittedAt: string
  details: string[]           // e.g. ["英語: 3コマ", "数学: 2コマ"]
  isUpdate: boolean
  captureElement?: HTMLElement | null  // optional element to screenshot
}

export async function downloadSubmissionReceiptPdf(params: SubmissionReceiptPdfParams): Promise<void> {
  const { sessionName, personName, personType, submittedAt, details, isUpdate, captureElement } = params

  let doc: jsPDF
  try {
    doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })
    await loadJapaneseFont(doc)
  } catch (err) {
    console.error('PDF生成に失敗しました:', err)
    return
  }

  const pageWidth = doc.internal.pageSize.getWidth()
  const centerX = pageWidth / 2
  const margin = 15

  // Title
  doc.setFontSize(18)
  doc.text(isUpdate ? '希望入力 更新記録' : '希望入力 提出記録', centerX, 25, { align: 'center' })

  // Content
  doc.setFontSize(11)
  const lines = [
    isUpdate ? '以下の内容で希望入力を更新しました。' : '以下の内容で希望入力を提出しました。',
    '',
    `■ セッション名: ${sessionName}`,
    `■ ${personType}名: ${personName}`,
    `■ 提出日時: ${submittedAt}`,
    '',
    '【提出内容】',
    ...details,
  ]

  let y = 38
  for (const line of lines) {
    doc.text(line, margin, y)
    y += 7
  }

  // Capture screenshot of the form if element provided
  if (captureElement) {
    try {
      const canvas = await html2canvas(captureElement, {
        scale: 1.5,
        useCORS: true,
        backgroundColor: '#ffffff',
        logging: false,
      })
      const imgWidth = pageWidth - margin * 2
      const imgHeight = (canvas.height / canvas.width) * imgWidth

      y += 4
      doc.setFontSize(10)
      doc.text('【入力画面スクリーンショット】', margin, y)
      y += 5

      // If screenshot fits on current page, add it; otherwise start new page
      const pageHeight = doc.internal.pageSize.getHeight()
      if (y + imgHeight > pageHeight - 20) {
        doc.addPage()
        y = 15
      }

      // Possibly span multiple pages
      let remainingHeight = imgHeight
      let srcY = 0
      while (remainingHeight > 0) {
        const availH = doc.internal.pageSize.getHeight() - y - 15
        const drawH = Math.min(remainingHeight, availH)
        // Compute source crop ratio
        const srcRatio = drawH / imgHeight
        const srcHeight = canvas.height * srcRatio

        // Create a cropped canvas for this page segment
        const cropCanvas = document.createElement('canvas')
        cropCanvas.width = canvas.width
        cropCanvas.height = Math.ceil(srcHeight)
        const ctx = cropCanvas.getContext('2d')
        if (ctx) {
          ctx.drawImage(canvas, 0, srcY, canvas.width, srcHeight, 0, 0, canvas.width, srcHeight)
          const cropImgData = cropCanvas.toDataURL('image/jpeg', 0.85)
          doc.addImage(cropImgData, 'JPEG', margin, y, imgWidth, drawH)
        }

        remainingHeight -= drawH
        srcY += srcHeight
        if (remainingHeight > 0) {
          doc.addPage()
          y = 15
        }
      }
    } catch (err) {
      console.error('スクリーンショット取得に失敗:', err)
      // Continue without screenshot
      y += 4
      doc.setFontSize(9)
      doc.setTextColor(120)
      doc.text('（スクリーンショットの取得に失敗しました）', margin, y)
      doc.setTextColor(0)
    }
  }

  // Footer on last page
  const lastPageH = doc.internal.pageSize.getHeight()
  doc.setFontSize(9)
  doc.setTextColor(120)
  doc.text(`発行日: ${new Date().toLocaleDateString('ja-JP')}`, centerX, lastPageH - 10, { align: 'center' })

  const label = isUpdate ? '更新記録' : '提出記録'
  const ts = submittedAt.replace(/[\/:]/g, '').replace(/\s+/g, '_')
  doc.save(`希望入力_${label}_${sessionName}_${personName}_${ts}.pdf`)
}
