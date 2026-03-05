/**
 * Constraint card system for auto-assign scoring.
 *
 * Each student can have ConstraintCardType[] that describe desired scheduling patterns.
 * Constraint cards have HIGHER priority than shared rules.
 *
 * Default cards (auto-enabled unless removed):
 *   - lateSlotNonExam: 受験生以外の後半コマ優先
 *   - groupContinuous: 集団後連続
 *
 * Conflict group (mutually exclusive per student):
 *   twoConsecutive / twoWithGap / oneSlotOnly / regularLink
 */
import type { Assignment, ConstraintCardType, RegularLesson, Student } from '../types'
import { getSlotNumber, getIsoDayOfWeek, getStudentSlotNumbersOnDate, getStudentSubjectsOnAdjacentSlots } from './assignments'

/** Labels for each constraint card type (UI display). */
export const CONSTRAINT_CARD_LABELS: Record<ConstraintCardType, string> = {
  lateSlotNonExam: '受験生以外の後半コマ優先',
  groupContinuous: '集団後連続',
  preferRegularTeacher: '通常講師優先',
  twoConsecutive: '2コマ連続',
  twoWithGap: '2コマ連続(一コマ空け)',
  oneSlotOnly: '一コマ限定',
  regularLink: '通常授業連結',
}

/** Short descriptions for each card type. */
export const CONSTRAINT_CARD_DESCRIPTIONS: Record<ConstraintCardType, string> = {
  lateSlotNonExam: '高3・中3以外は3限以降に配置しやすくし、2人ペアの形成を促進',
  groupContinuous: '集団授業がある日の中3は、午前の後に早いコマから2コマ連続で配置',
  preferRegularTeacher: '通常授業の講師を優先して配置',
  twoConsecutive: '生徒を2コマ連続で配置する（複数科目の残コマがある場合、科目は前後のコマで分ける）',
  twoWithGap: '生徒を2コマ連続で配置するが、間に1コマ入れる（複数科目の残コマがある場合、科目は前後のコマで分ける）',
  oneSlotOnly: '生徒を1日1コマに限定する',
  regularLink: '通常授業の前後に特別講習のコマをつなげ2コマ連続とする（複数科目の残コマがある場合、科目は前後のコマで分ける）',
}

/** Cards enabled by default when creating new sessions. */
export const DEFAULT_CONSTRAINT_CARDS: ConstraintCardType[] = ['lateSlotNonExam', 'groupContinuous']

/** All constraint card types in display order. */
export const ALL_CONSTRAINT_CARDS: ConstraintCardType[] = [
  'lateSlotNonExam',
  'groupContinuous',
  'preferRegularTeacher',
  'twoConsecutive',
  'twoWithGap',
  'oneSlotOnly',
  'regularLink',
]

/** Conflict group: these cards are mutually exclusive for a single student. */
export const CONSTRAINT_CARD_CONFLICT_GROUP: ConstraintCardType[] = [
  'twoConsecutive', 'twoWithGap', 'oneSlotOnly', 'regularLink',
]

/** Check if a grade is "exam grade" (受験生): 高3 or 中3 */
export const isExamGrade = (grade: string): boolean => grade === '高3' || grade === '中3'

/**
 * Validate that a set of constraint cards have no conflicts.
 * Returns warning messages (empty = no conflicts).
 */
export const validateConstraintCards = (cards: ConstraintCardType[]): string[] => {
  const warnings: string[] = []
  const conflicting = cards.filter((c) => CONSTRAINT_CARD_CONFLICT_GROUP.includes(c))
  if (conflicting.length > 1) {
    const labels = conflicting.map((c) => CONSTRAINT_CARD_LABELS[c]).join('、')
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
 * @param regularLessons Regular lessons (for 'regularLink' and 'preferRegularTeacher')
 * @param groupLessons   Group lessons (for 'groupContinuous')
 * @param teacherId      The teacher being considered (for 'preferRegularTeacher')
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
  const cards = student.constraintCards ?? DEFAULT_CONSTRAINT_CARDS
  if (cards.length === 0) return { score: 0, blocked: false }

  let score = 0
  const date = candidateSlot.split('_')[0]
  const slotNum = getSlotNumber(candidateSlot)
  const existingSlotNums = getStudentSlotNumbersOnDate(assignments, student.id, date)
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
          return { score: -99999, blocked: true, blockReason: '集団後連続: 既に2コマ配置済み' }
        }
        break
      }

      case 'preferRegularTeacher': {
        // Bonus if this teacher is the student's regular teacher
        if (teacherId) {
          const isRegular = regularLessons.some(
            (rl) => rl.teacherId === teacherId && rl.studentIds.includes(student.id),
          )
          if (isRegular) {
            score += 1500
          }
        }
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
          return { score: -99999, blocked: true, blockReason: '一コマ限定: 既に1コマ配置済み' }
        }
        score += 100 // Small bonus for being placed (since no more can be added)
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
    }
  }

  return { score, blocked: false }
}