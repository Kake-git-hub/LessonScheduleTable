import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
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
    // Try multiple CDN sources for Japanese TTF fonts (jsPDF needs TrueType, not CFF-based OTF)
    const urls = [
      'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf',
      'https://raw.githubusercontent.com/google/fonts/main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf',
      'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/kosugimaru/KosugiMaru-Regular.ttf',
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
  doc.setFont('NotoSansJP')
}

// ---------- Schedule PDF (A3 landscape, one week per page) ----------

export type SchedulePdfParams = {
  sessionName: string
  startDate: string
  endDate: string
  slotsPerDay: number
  holidays: string[]
  assignments: Record<string, Assignment[]>
  getTeacherName: (id: string) => string
  getStudentName: (id: string) => string
  getStudentSubject: (a: Assignment, studentId: string) => string
  getIsoDayOfWeek: (date: string) => number
}

export async function exportSchedulePdf(params: SchedulePdfParams): Promise<void> {
  const {
    sessionName, startDate, endDate, slotsPerDay, holidays,
    assignments, getTeacherName, getStudentName, getStudentSubject, getIsoDayOfWeek
  } = params

  if (!startDate || !endDate) { alert('開始日・終了日を設定してください。'); return }

  const holidaySet = new Set(holidays)
  const dayNames = ['日', '月', '火', '水', '木', '金', '土']

  // Build all dates
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

  // Group dates into weeks (Mon–Sun)
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

  // Create PDF - A3 landscape
  let doc: jsPDF
  try {
    doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a3' })
    await loadJapaneseFont(doc)
  } catch (err) {
    alert('PDF生成に失敗しました: ' + String(err))
    return
  }

  const dowOrder = [1, 2, 3, 4, 5, 6, 0] // Mon-Sun

  for (let wi = 0; wi < weeks.length; wi++) {
    if (wi > 0) doc.addPage('a3', 'landscape')

    const weekDates = weeks[wi]
    const firstDate = weekDates[0]

    // Pad to full Mon–Sun week
    const fullWeek: (string | null)[] = []
    const firstDow = getIsoDayOfWeek(firstDate)
    const startPad = firstDow === 0 ? 6 : firstDow - 1
    for (let p = 0; p < startPad; p++) fullWeek.push(null)
    for (const d of weekDates) fullWeek.push(d)
    while (fullWeek.length < 7) fullWeek.push(null)

    // Build cell data for each slot
    const getCellParts = (slotKey: string): { teacher: string; student1: string; student2: string }[] => {
      const slotAssignments = assignments[slotKey] ?? []
      if (slotAssignments.length === 0) return []
      return slotAssignments.map((a) => {
        const teacher = getTeacherName(a.teacherId)
        const students = a.studentIds.map((sid) => {
          const name = getStudentName(sid)
          const subj = getStudentSubject(a, sid)
          return `${name}(${subj})`
        })
        return { teacher, student1: students[0] ?? '', student2: students[1] ?? '' }
      })
    }

    // Title
    const [, fm, fd] = firstDate.split('-')
    const lastDate = weekDates[weekDates.length - 1]
    const [, lm, ld] = lastDate.split('-')
    doc.setFontSize(14)
    doc.text(`${sessionName}  ${Number(fm)}/${Number(fd)} - ${Number(lm)}/${Number(ld)}`, 10, 12)

    // Build header: 2 rows
    // Row 1: empty | date spans (each spanning 3 cols: 講師, 生徒, 生徒)
    // Row 2: empty | 講師, 生徒, 生徒 × 7
    const headerRow1: string[] = ['']
    const headerRow2: string[] = ['']
    for (let i = 0; i < 7; i++) {
      const date = fullWeek[i]
      if (!date) {
        headerRow1.push(dayNames[dowOrder[i]], '', '')
      } else {
        const [, mm, dd] = date.split('-')
        headerRow1.push(`${Number(mm)}/${Number(dd)}(${dayNames[dowOrder[i]]})`, '', '')
      }
      headerRow2.push('講師', '生徒', '生徒')
    }

    // Build data rows - track which slot each row belongs to for alternating colors
    const bodyRows: string[][] = []
    const rowSlotNum: number[] = [] // track slot number per row for coloring
    for (let s = 1; s <= slotsPerDay; s++) {
      // Find max number of assignment pairs across all days for this slot
      let maxPairs = 1
      for (let i = 0; i < 7; i++) {
        const date = fullWeek[i]
        if (date && !holidaySet.has(date)) {
          const parts = getCellParts(`${date}_${s}`)
          if (parts.length > maxPairs) maxPairs = parts.length
        }
      }

      for (let pairIdx = 0; pairIdx < maxPairs; pairIdx++) {
        const row: string[] = [pairIdx === 0 ? `${s}限` : '']
        for (let i = 0; i < 7; i++) {
          const date = fullWeek[i]
          if (!date) {
            row.push('', '', '')
          } else if (holidaySet.has(date)) {
            row.push(pairIdx === 0 ? '休' : '', '', '')
          } else {
            const parts = getCellParts(`${date}_${s}`)
            const pair = parts[pairIdx]
            if (pair) {
              row.push(pair.teacher, pair.student1, pair.student2)
            } else {
              row.push('', '', '')
            }
          }
        }
        bodyRows.push(row)
        rowSlotNum.push(s)
      }
    }

    // Use autoTable for rendering
    const pageWidth = doc.internal.pageSize.getWidth()
    const margin = 6
    const availWidth = pageWidth - margin * 2
    const slotColWidth = 10
    const dayColWidth = (availWidth - slotColWidth) / 21 // 7 days × 3 cols

    // Column styles
    const columnStyles: Record<number, { cellWidth: number; halign?: 'center' | 'left' }> = { 0: { cellWidth: slotColWidth, halign: 'center' } }
    for (let c = 1; c <= 21; c++) {
      columnStyles[c] = { cellWidth: dayColWidth }
    }

    autoTable(doc, {
      startY: 16,
      margin: { left: margin, right: margin },
      theme: 'grid',
      styles: {
        font: 'NotoSansJP',
        fontSize: 7,
        cellPadding: 1,
        lineWidth: 0.2,
        lineColor: [80, 80, 80],
        valign: 'middle',
        overflow: 'linebreak',
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 0.3,
        lineColor: [40, 40, 40],
      },
      head: [headerRow1, headerRow2],
      body: bodyRows,
      columnStyles,
      didParseCell: (hookData) => {
        // Merge date header cells (row 0: span 3 cols per day)
        if (hookData.section === 'head' && hookData.row.index === 0) {
          const col = hookData.column.index
          if (col === 0) {
            // empty cell
          } else if ((col - 1) % 3 === 0) {
            // First of the 3 day-columns → span
            hookData.cell.colSpan = 3
            hookData.cell.styles.halign = 'center'
            hookData.cell.styles.fontStyle = 'bold'
          }
        }
        // Slot label column bold
        if (hookData.section === 'body' && hookData.column.index === 0 && hookData.cell.text.join('')) {
          hookData.cell.styles.fontStyle = 'bold'
          hookData.cell.styles.halign = 'center'
        }
        // Alternating slot group colors for readability
        if (hookData.section === 'body') {
          const slotNum = rowSlotNum[hookData.row.index] ?? 1
          if (slotNum % 2 === 0) {
            hookData.cell.styles.fillColor = [245, 247, 250]
          }
        }
      },
    })
  }

  try {
    doc.save(`コマ割り_${sessionName}.pdf`)
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
