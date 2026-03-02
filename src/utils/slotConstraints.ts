/**
 * Slot constraint evaluation for auto-assign scoring.
 *
 * Each student can have SlotConstraint cards that describe desired scheduling patterns
 * (e.g. "2コマ連続", "1コマ空けて2コマ").
 *
 * - 'must' constraints: applied as hard filters (block invalid placements) AND a large scoring bonus
 * - 'prefer' constraints: applied only as scoring bonuses/penalties
 */
import type { Assignment, SlotConstraint, Student } from '../types'
import { getSlotNumber, getStudentSlotNumbersOnDate, getStudentSubjectsOnAdjacentSlots } from './assignments'

export interface ConstraintEvalResult {
  /** Total score adjustment from all constraints */
  score: number
  /** Descriptions of violated 'must' constraints (empty if all OK) */
  mustViolations: string[]
}

/**
 * Labels for constraint types (UI display).
 */
export const SLOT_CONSTRAINT_LABELS: Record<string, string> = {
  'consecutive': '連続コマ',
  'gap-then-consecutive': '空けて連続',
}

/**
 * Default params for each constraint type.
 */
export const defaultConstraintParams = (type: string): SlotConstraint['params'] => {
  switch (type) {
    case 'consecutive':
      return { count: 2, diffSubject: false }
    case 'gap-then-consecutive':
      return { count: 2, gapSlots: 1, diffSubject: false }
    default:
      return {}
  }
}

/**
 * Evaluate all slot constraints for a student being placed in a candidate slot.
 *
 * @param student       The student being evaluated
 * @param candidateSlot The slot key being considered (e.g. "2026-03-15_2")
 * @param assignments   Current assignment state
 * @param slotsPerDay   Total slots per day (from session settings)
 * @returns Score adjustment and any must-violations
 */
export const evaluateSlotConstraints = (
  student: Student,
  candidateSlot: string,
  assignments: Record<string, Assignment[]>,
  slotsPerDay: number,
): ConstraintEvalResult => {
  const constraints = student.slotConstraints ?? []
  if (constraints.length === 0) return { score: 0, mustViolations: [] }

  let score = 0
  const mustViolations: string[] = []
  const date = candidateSlot.split('_')[0]
  const slotNum = getSlotNumber(candidateSlot)

  // Student's already-assigned slot numbers on this date
  const existingSlotNums = getStudentSlotNumbersOnDate(assignments, student.id, date)

  for (const c of constraints) {
    const weight = c.priority === 'must' ? 150 : 40

    switch (c.type) {
      case 'consecutive': {
        // Desired: student has N consecutive slots on a single day
        const count = c.params.count ?? 2
        const result = evaluateConsecutive(slotNum, existingSlotNums, count, slotsPerDay)
        if (result.satisfied) {
          score += weight
          // Bonus for different subjects if requested
          if (c.params.diffSubject && result.adjacentSlotNum !== undefined) {
            const adjSubjects = getStudentSubjectsOnAdjacentSlots(assignments, student.id, date, slotNum)
            if (adjSubjects.length > 0) {
              // Having adjacent slots with subjects is what we want — the actual subject choice
              // is handled by autoAssign's subject selection. Just add a small bonus here.
              score += 10
            }
          }
        } else {
          // Not consecutive — penalty or violation
          if (existingSlotNums.length > 0) {
            // Already have slots but not consecutive — bad
            score -= weight
            if (c.priority === 'must') {
              mustViolations.push(`${count}コマ連続が必要`)
            }
          } else {
            // First slot on this date — check if forming a consecutive block is still possible
            if (canFormConsecutiveBlock(slotNum, count, slotsPerDay)) {
              score += weight / 4 // Small bonus: good starting position
            }
          }
        }
        break
      }

      case 'gap-then-consecutive': {
        // Desired: gap of M slots, then N consecutive slots
        // e.g. gapSlots=1, count=2 → slot 1, skip 2, slots 3-4
        const gapSlots = c.params.gapSlots ?? 1
        const count = c.params.count ?? 2
        const result = evaluateGapThenConsecutive(slotNum, existingSlotNums, gapSlots, count, slotsPerDay)
        if (result.satisfied) {
          score += weight
          if (c.params.diffSubject) {
            const adjSubjects = getStudentSubjectsOnAdjacentSlots(assignments, student.id, date, slotNum)
            if (adjSubjects.length > 0) score += 10
          }
        } else {
          if (existingSlotNums.length > 0) {
            score -= weight
            if (c.priority === 'must') {
              mustViolations.push(`${gapSlots}コマ空けて${count}コマ連続が必要`)
            }
          } else {
            if (canFormGapConsecutiveBlock(slotNum, gapSlots, count, slotsPerDay)) {
              score += weight / 4
            }
          }
        }
        break
      }
    }
  }

  return { score, mustViolations }
}

// ── Internal helpers ──────────────────────────────────────

interface ConsecutiveResult {
  satisfied: boolean
  adjacentSlotNum?: number
}

/**
 * Check if placing at slotNum forms (or extends) a consecutive block of `count` slots.
 */
const evaluateConsecutive = (
  slotNum: number,
  existingSlots: number[],
  count: number,
  _slotsPerDay: number,
): ConsecutiveResult => {
  if (existingSlots.length === 0) {
    // First slot — can't be consecutive yet, but is OK
    return { satisfied: false }
  }

  // Build the set including the candidate
  const allSlots = [...existingSlots, slotNum].sort((a, b) => a - b)

  // Check if there's a consecutive run of `count` that includes slotNum
  for (let start = 0; start <= allSlots.length - count; start++) {
    let isRun = true
    for (let i = 1; i < count; i++) {
      if (allSlots[start + i] !== allSlots[start] + i) { isRun = false; break }
    }
    if (isRun) {
      const runSlots = allSlots.slice(start, start + count)
      if (runSlots.includes(slotNum)) {
        // Find an adjacent existing slot
        const adj = existingSlots.find((n) => Math.abs(n - slotNum) === 1)
        return { satisfied: true, adjacentSlotNum: adj }
      }
    }
  }

  // Even without a full run, being adjacent to existing is partially good
  const isAdj = existingSlots.some((n) => Math.abs(n - slotNum) === 1)
  if (isAdj && existingSlots.length + 1 < count) {
    // Building towards the target — partial satisfaction
    return { satisfied: false, adjacentSlotNum: existingSlots.find((n) => Math.abs(n - slotNum) === 1) }
  }

  return { satisfied: false }
}

/**
 * Check if slotNum can be the start of a consecutive block of `count` within the day.
 */
const canFormConsecutiveBlock = (slotNum: number, count: number, slotsPerDay: number): boolean => {
  // Check if there's room for a block starting at or before slotNum
  for (let start = Math.max(1, slotNum - count + 1); start <= slotNum; start++) {
    if (start + count - 1 <= slotsPerDay) return true
  }
  return false
}

/**
 * Evaluate gap-then-consecutive pattern.
 * Pattern: [existing slot] → [gapSlots empty] → [count consecutive slots]
 * OR: [count consecutive slots] → [gapSlots empty] → [existing slot]
 */
const evaluateGapThenConsecutive = (
  slotNum: number,
  existingSlots: number[],
  gapSlots: number,
  count: number,
  slotsPerDay: number,
): ConsecutiveResult => {
  if (existingSlots.length === 0) return { satisfied: false }

  const allSlots = [...existingSlots, slotNum].sort((a, b) => a - b)

  // Look for the pattern: a group of isolated slot(s), then exactly `gapSlots` gap, then `count` consecutive
  // We need to find if the candidate slot fits into such a pattern

  // Check all possible consecutive blocks that include slotNum
  for (let blockStart = Math.max(1, slotNum - count + 1); blockStart <= slotNum; blockStart++) {
    const blockEnd = blockStart + count - 1
    if (blockEnd > slotsPerDay) continue

    // Check that all positions in the block are filled
    const blockSlots = Array.from({ length: count }, (_, i) => blockStart + i)
    if (!blockSlots.every((s) => allSlots.includes(s))) continue

    // Check for a slot exactly gapSlots before the block or after the block
    const slotBeforeGap = blockStart - gapSlots - 1
    const slotAfterGap = blockEnd + gapSlots + 1

    if ((slotBeforeGap >= 1 && allSlots.includes(slotBeforeGap)) ||
        (slotAfterGap <= slotsPerDay && allSlots.includes(slotAfterGap))) {
      return { satisfied: true }
    }
  }

  // Check if we're building towards the pattern
  const isAdj = existingSlots.some((n) => Math.abs(n - slotNum) === 1)
  if (isAdj) return { satisfied: false }

  // Check if this slot is exactly gapSlots away from an existing slot
  const isGapAway = existingSlots.some((n) => Math.abs(n - slotNum) === gapSlots + 1)
  if (isGapAway) return { satisfied: false }

  return { satisfied: false }
}

/**
 * Check if a gap-then-consecutive block can still be formed starting from/around slotNum.
 */
const canFormGapConsecutiveBlock = (
  slotNum: number,
  gapSlots: number,
  count: number,
  slotsPerDay: number,
): boolean => {
  const totalNeeded = 1 + gapSlots + count // e.g. 1 + 1 + 2 = 4 slots span
  // Check if there's enough room in either direction
  return (slotNum + totalNeeded - 1 <= slotsPerDay) || (slotNum - totalNeeded + 1 >= 1)
}

/**
 * Summarize a student's constraints as a short display string.
 */
export const summarizeConstraints = (constraints: SlotConstraint[]): string => {
  return constraints.map((c) => {
    const priority = c.priority === 'must' ? '必須' : '希望'
    switch (c.type) {
      case 'consecutive':
        return `${c.params.count ?? 2}コマ連続${c.params.diffSubject ? '(別教科)' : ''}[${priority}]`
      case 'gap-then-consecutive':
        return `${c.params.gapSlots ?? 1}コマ空け→${c.params.count ?? 2}コマ連続${c.params.diffSubject ? '(別教科)' : ''}[${priority}]`
      default:
        return `${c.type}[${priority}]`
    }
  }).join(', ')
}

/**
 * Validate that a set of constraints don't conflict with each other.
 * Returns an array of warning messages (empty = no conflicts).
 */
export const validateConstraints = (constraints: SlotConstraint[]): string[] => {
  const warnings: string[] = []
  if (constraints.length <= 1) return warnings

  const types = constraints.map((c) => c.type)
  const hasConsecutive = types.includes('consecutive')
  const hasGapConsecutive = types.includes('gap-then-consecutive')

  // consecutive + gap-then-consecutive conflict: both dictate placement pattern
  if (hasConsecutive && hasGapConsecutive) {
    warnings.push('「連続コマ」と「空けて連続」は同時に設定できません。どちらか一方にしてください。')
  }

  // Multiple of the same type
  const consecutiveCards = constraints.filter((c) => c.type === 'consecutive')
  if (consecutiveCards.length > 1) {
    const counts = consecutiveCards.map((c) => c.params.count ?? 2)
    if (new Set(counts).size > 1) {
      warnings.push('複数の「連続コマ」制約のコマ数が異なります。1つにまとめてください。')
    } else {
      warnings.push('「連続コマ」制約が重複しています。1つにまとめてください。')
    }
  }

  const gapCards = constraints.filter((c) => c.type === 'gap-then-consecutive')
  if (gapCards.length > 1) {
    warnings.push('「空けて連続」制約が重複しています。1つにまとめてください。')
  }

  return warnings
}
