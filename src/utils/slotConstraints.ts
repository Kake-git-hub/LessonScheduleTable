/**
 * Constraint card system for auto-assign scoring.
 *
 * Each student can have ConstraintCardType[] that describe desired scheduling patterns.
 * Constraint cards have HIGHER priority than shared rules.
 *
 * Default cards (auto-enabled unless removed):
 *   - twoSlotLimit: 2コマ上限
 *   - lateSlotNonExam: 受験生以外の後半コマ優先
 *   - groupContinuous: 集団後2コマ連続
 *
 * Conflict groups (mutually exclusive per student):
 *   pattern: twoConsecutive / twoWithGap / groupContinuous / regularLink
 *   daily-limit: oneSlotOnly / twoSlotLimit / threeSlotLimit
 *   one-slot: oneSlotOnly / twoConsecutive / twoWithGap / groupContinuous / regularLink
 */
import type { Assignment, ConstraintCardType, RegularLesson, Student } from '../types'
import { getSlotNumber, getIsoDayOfWeek, getStudentNonGroupSlotNumbersOnDate, getStudentSubjectsOnAdjacentSlots } from './assignments'

/** Labels for each constraint card type (UI display). */
export const CONSTRAINT_CARD_LABELS: Record<ConstraintCardType, string> = {
  oneSlotOnly: '1コマ上限',
  twoSlotLimit: '2コマ上限',
  threeSlotLimit: '3コマ上限',
  twoConsecutive: '2コマ連続',
  twoWithGap: '2コマ連続(一コマ空け)',
  groupContinuous: '集団後2コマ連続',
  regularLink: '通常授業連結',
  forceRegularTeacher: '通常講師強制',
  priorityAssign: '優先割振',
  lateSlotNonExam: '受験生以外の後半コマ優先',
  earlySlotPreference: '小学生は2限寄り',
  lateSlotPreference: '中高生は5限寄り',
  avoidSlot1: '1限回避',
}

/** Short descriptions for each card type. */
export const CONSTRAINT_CARD_DESCRIPTIONS: Record<ConstraintCardType, string> = {
  oneSlotOnly: '[絶対] 生徒を1日1コマに限定する。集団授業はこの上限に含めない',
  twoSlotLimit: '[絶対] 生徒を1日2コマまでに制限する（デフォルト）。集団授業はこの上限に含めない',
  threeSlotLimit: '[絶対] 生徒を1日3コマまでに制限する。集団授業はこの上限に含めない',
  twoConsecutive: '[絶対] 生徒を2コマ連続で配置する（複数科目の残コマがある場合、科目は前後のコマで分ける）',
  twoWithGap: '[絶対] 生徒を2コマ連続で配置するが、間に1コマ入れる（複数科目の残コマがある場合、科目は前後のコマで分ける）',
  groupContinuous: '[絶対] 集団授業がある日の中3は、午前の後に早いコマから2コマ連続で配置',
  regularLink: '[絶対] 通常授業の前後に特別講習のコマをつなげ2コマ連続とする（複数科目の残コマがある場合、科目は前後のコマで分ける）',
  forceRegularTeacher: '[絶対] 通常授業の講師以外への配置を不可にする',
  priorityAssign: '[推奨] 他の生徒より優先して割り振る（中3デフォルト）',
  lateSlotNonExam: '[推奨] 高3・中3以外は3限以降に配置しやすくし、2人ペアの形成を促進',
  earlySlotPreference: '[推奨] 小学生を2限優先で配置しやすくする（2限 > 3限 > 4限 > 5限、1限はなるべく避ける）',
  lateSlotPreference: '[推奨] 中高生を5限以降優先で配置しやすくする（5限以降 > 4限 > 3限 > 2限 > 1限）',
  avoidSlot1: '[絶対] 1限への配置を禁止する',
}

/** Get default constraint cards for a student based on grade.
 *  - lateSlotNonExam: default for non-exam grades (not 中3/高3)
 *  - groupContinuous: default for 中3 only
 */
export const getDefaultConstraintCards = (grade: string): ConstraintCardType[] => {
  const cards: ConstraintCardType[] = ['twoSlotLimit']
  if (!isExamGrade(grade)) {
    cards.push('lateSlotNonExam')
  }
  if (grade === '中3') {
    cards.push('groupContinuous')
    cards.push('priorityAssign')
  }
  // Slot preference defaults
  if (grade.startsWith('小')) {
    cards.push('earlySlotPreference')
  } else {
    cards.push('lateSlotPreference')
  }
  return cards
}

/** Static fallback (used when grade is unknown). */
export const DEFAULT_CONSTRAINT_CARDS: ConstraintCardType[] = ['twoSlotLimit', 'lateSlotNonExam', 'groupContinuous', 'priorityAssign', 'lateSlotPreference']

/** All constraint card types in display order. */
export const ALL_CONSTRAINT_CARDS: ConstraintCardType[] = [
  'oneSlotOnly',
  'twoSlotLimit',
  'threeSlotLimit',
  'twoConsecutive',
  'twoWithGap',
  'groupContinuous',
  'regularLink',
  'forceRegularTeacher',
  'priorityAssign',
  'lateSlotNonExam',
  'earlySlotPreference',
  'lateSlotPreference',
  'avoidSlot1',
]

/** Conflict group: scheduling pattern cards are mutually exclusive for a single student. */
export const CONSTRAINT_CARD_CONFLICT_GROUP: ConstraintCardType[] = [
  'twoConsecutive', 'twoWithGap', 'groupContinuous', 'regularLink',
]

/** Daily limit conflict group: daily slot limit cards are mutually exclusive. */
export const DAILY_LIMIT_CONFLICT_GROUP: ConstraintCardType[] = [
  'oneSlotOnly', 'twoSlotLimit', 'threeSlotLimit',
]

/** One-slot cap conflicts with any card that requires or strongly prefers 2 connected slots. */
export const ONE_SLOT_CONFLICT_GROUP: ConstraintCardType[] = [
  'oneSlotOnly', 'twoConsecutive', 'twoWithGap', 'groupContinuous', 'regularLink',
]

/** Slot preference conflict group: early vs late preference are mutually exclusive. */
export const SLOT_PREFERENCE_CONFLICT_GROUP: ConstraintCardType[] = [
  'earlySlotPreference', 'lateSlotPreference',
]

export const CONSTRAINT_CARD_CONFLICT_GROUPS: ConstraintCardType[][] = [
  CONSTRAINT_CARD_CONFLICT_GROUP,
  DAILY_LIMIT_CONFLICT_GROUP,
  ONE_SLOT_CONFLICT_GROUP,
  SLOT_PREFERENCE_CONFLICT_GROUP,
]

/** Check if a grade is "exam grade" (受験生): 高3 or 中3 */
export const isExamGrade = (grade: string): boolean => grade === '高3' || grade === '中3'

/**
 * Validate that a set of constraint cards have no conflicts.
 * Returns warning messages (empty = no conflicts).
 */
export const validateConstraintCards = (cards: ConstraintCardType[]): string[] => {
  const warnings: string[] = []
  const seen = new Set<string>()
  for (const group of CONSTRAINT_CARD_CONFLICT_GROUPS) {
    const conflicting = cards.filter((c) => group.includes(c))
    if (conflicting.length <= 1) continue
    const labels = conflicting.map((c) => CONSTRAINT_CARD_LABELS[c]).join('、')
    if (seen.has(labels)) continue
    seen.add(labels)
    warnings.push(`${labels} は競合しています。どれか1つだけ選択してください。`)
  }
  return warnings
}

/**
 * Summarize constraint cards as a short display string.
 */
export const summarizeConstraintCards = (cards: ConstraintCardType[]): string => {
  if (cards.length === 0) return ''
  return cards.map((c) => CONSTRAINT_CARD_LABELS[c]).join(', ')
}

export interface ConstraintEvalResult {
  /** Total score adjustment from all constraints */
  score: number
  /** Whether placement is blocked (hard constraint violation) */
  blocked: boolean
  /** Reason for blocking */
  blockReason?: string
}

/**
 * Evaluate constraint cards for a student being placed in a candidate slot.
 *
 * @param student        The student being evaluated
 * @param candidateSlot  The slot key being considered (e.g. "2026-03-15_2")
 * @param assignments    Current assignment state
 * @param slotsPerDay    Total slots per day (from session settings)
 * @param regularLessons Regular lessons (for 'regularLink' and 'forceRegularTeacher')
 * @param groupLessons   Group lessons (for 'groupContinuous')
 * @param teacherId      The teacher being considered (for 'forceRegularTeacher')
 */
export const evaluateConstraintCards = (
  student: Student,
  candidateSlot: string,
  assignments: Record<string, Assignment[]>,
  slotsPerDay: number,
  regularLessons: RegularLesson[],
  groupLessons: Array<{ studentIds: string[]; dayOfWeek: number; slotNumber: number }>,
  teacherId?: string,
): ConstraintEvalResult => {
  const cards = student.constraintCards ?? getDefaultConstraintCards(student.grade)
  if (cards.length === 0) return { score: 0, blocked: false }

  let score = 0
  const date = candidateSlot.split('_')[0]
  const slotNum = getSlotNumber(candidateSlot)
  const existingSlotNums = getStudentNonGroupSlotNumbersOnDate(assignments, student.id, date)
  const dayOfWeek = getIsoDayOfWeek(date)

  for (const card of cards) {
    switch (card) {
      case 'lateSlotNonExam': {
        // Non-exam students get bonus for slots 3+, penalty for early slots
        if (!isExamGrade(student.grade)) {
          if (slotNum >= 3) {
            score += 800
          } else if (slotNum <= 1) {
            score -= 400
          }
        }
        break
      }

      case 'groupContinuous': {
        // 中3 students with group lesson on this date → early consecutive from slot 1 or 2
        if (student.grade !== '中3') break
        const hasGroupOnDate = groupLessons.some(
          (gl) => gl.dayOfWeek === dayOfWeek && gl.studentIds.includes(student.id),
        )
        if (!hasGroupOnDate) break

        // Strong preference for slots 1 and 2
        if (slotNum === 1 || slotNum === 2) {
          score += 2000
        } else if (slotNum === 3) {
          score += 500
        } else {
          score -= 1000
        }

        // Consecutive bonus
        if (existingSlotNums.length > 0 && existingSlotNums.length < 2) {
          const isConsecutive = existingSlotNums.some((n) => Math.abs(n - slotNum) === 1)
          score += isConsecutive ? 2000 : -1500
        }
        // Already has 2+ individual slots → block
        if (existingSlotNums.length >= 2) {
          return { score: -99999, blocked: true, blockReason: '集団後2コマ連続: 既に2コマ配置済み' }
        }
        break
      }

      case 'forceRegularTeacher': {
        // Hard block: only allow the student's regular teacher
        if (teacherId) {
          const hasRegularLesson = regularLessons.some((rl) => rl.studentIds.includes(student.id))
          if (hasRegularLesson) {
            const isRegular = regularLessons.some(
              (rl) => rl.teacherId === teacherId && rl.studentIds.includes(student.id),
            )
            if (!isRegular) {
              return { score: -99999, blocked: true, blockReason: '通常講師強制: 通常授業の講師ではありません' }
            }
            score += 1500
          }
        }
        break
      }

      case 'priorityAssign': {
        // Large bonus to prioritize this student over others
        score += 3000
        break
      }

      case 'twoConsecutive': {
        // Student should be in 2 consecutive slots per day
        // Hard limit: max 2 slots per day
        if (existingSlotNums.length >= 2) {
          return { score: -99999, blocked: true, blockReason: '2コマ連続: 既に2コマ配置済み' }
        }

        if (existingSlotNums.length === 1) {
          // Must be consecutive to existing slot
          const isConsecutive = existingSlotNums.some((n) => Math.abs(n - slotNum) === 1)
          if (isConsecutive) {
            score += 3000
            // Bonus for different subject in adjacent slots
            const adjSubjects = getStudentSubjectsOnAdjacentSlots(assignments, student.id, date, slotNum)
            if (adjSubjects.length > 0) {
              score += 200 // Having subjects means we can differentiate
            }
          } else {
            return { score: -99999, blocked: true, blockReason: '2コマ連続: 連続していないコマ' }
          }
        } else {
          // First slot — prefer positions where consecutive is still possible
          if (slotNum < slotsPerDay) {
            score += 500 // Can still add consecutive after
          }
        }
        break
      }

      case 'twoWithGap': {
        // Student should be in 2 slots with exactly 1 gap between them
        if (existingSlotNums.length >= 2) {
          return { score: -99999, blocked: true, blockReason: '2コマ連続(一コマ空け): 既に2コマ配置済み' }
        }

        if (existingSlotNums.length === 1) {
          const existing = existingSlotNums[0]
          const gap = Math.abs(slotNum - existing)
          if (gap === 2) {
            score += 3000 // Exactly 1 slot gap
            const adjSubjects = getStudentSubjectsOnAdjacentSlots(assignments, student.id, date, slotNum)
            if (adjSubjects.length > 0) score += 200
          } else {
            return { score: -99999, blocked: true, blockReason: '2コマ連続(一コマ空け): 間隔が正しくありません' }
          }
        } else {
          // First slot — check if gap pattern is still possible
          if (slotNum + 2 <= slotsPerDay || slotNum - 2 >= 1) {
            score += 500
          }
        }
        break
      }

      case 'oneSlotOnly': {
        // Hard limit: 1 slot per day
        if (existingSlotNums.length >= 1) {
          return { score: -99999, blocked: true, blockReason: '1コマ上限: 既に1コマ配置済み' }
        }
        score += 100 // Small bonus for being placed (since no more can be added)
        break
      }

      case 'twoSlotLimit': {
        // Hard limit: 2 slots per day (default behavior)
        if (existingSlotNums.length >= 2) {
          return { score: -99999, blocked: true, blockReason: '2コマ上限: 既に2コマ配置済み' }
        }
        break
      }

      case 'threeSlotLimit': {
        // Hard limit: 3 slots per day
        if (existingSlotNums.length >= 3) {
          return { score: -99999, blocked: true, blockReason: '3コマ上限: 既に3コマ配置済み' }
        }
        break
      }

      case 'regularLink': {
        // Link special lesson adjacent to regular lesson on this day-of-week
        const studentRegularSlots = regularLessons
          .filter((rl) => rl.dayOfWeek === dayOfWeek && rl.studentIds.includes(student.id))
          .map((rl) => rl.slotNumber)

        if (studentRegularSlots.length === 0) break // No regular on this day — N/A

        // Max 1 special slot per day for regularLink (forming 2コマ: regular + special)
        if (existingSlotNums.length >= 1) {
          return { score: -99999, blocked: true, blockReason: '通常授業連結: 既に特別講習コマ配置済み' }
        }

        const isAdjacent = studentRegularSlots.some((n) => Math.abs(n - slotNum) === 1)
        if (isAdjacent) {
          score += 3000
          const adjSubjects = getStudentSubjectsOnAdjacentSlots(assignments, student.id, date, slotNum)
          if (adjSubjects.length > 0) score += 200
        } else {
          score -= 2000 // Not adjacent to regular → strong penalty
        }
        break
      }

      case 'earlySlotPreference': {
        // 小学生は 2限 > 3限 > 4限 > 5限 を優先し、1限はできるだけ避ける
        if (slotNum === 2) {
          score += 1400
        } else if (slotNum === 3) {
          score += 650
        } else if (slotNum === 4) {
          score += 150
        } else if (slotNum >= 5) {
          score -= 350 * (slotNum - 4)
        } else if (slotNum === 1) {
          score -= 900
        }
        break
      }

      case 'lateSlotPreference': {
        // 中高生は 5限以降 > 4限 > 3限 > 2限 > 1限 を強めに優先
        if (slotNum >= 5) {
          score += 1400 + 100 * (slotNum - 5)
        } else if (slotNum === 4) {
          score += 450
        } else if (slotNum === 3) {
          score -= 150
        } else if (slotNum === 2) {
          score -= 700
        } else {
          score -= 1300
        }
        break
      }

      case 'avoidSlot1': {
        // 1限回避: ハード制約 — 1限への配置を完全にブロック
        if (slotNum === 1) {
          return { score: -99999, blocked: true, blockReason: '1限回避: 1限への配置は禁止されています' }
        }
        break;
      }
    }
  }

  return { score, blocked: false }
}