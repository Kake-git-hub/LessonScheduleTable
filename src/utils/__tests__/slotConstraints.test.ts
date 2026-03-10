import { describe, expect, it } from 'vitest'

import type { Student } from '../../types'
import { evaluateConstraintCards } from '../slotConstraints'

const makeStudent = (grade: string, constraintCards: Student['constraintCards']): Student => ({
  id: 's1',
  name: 'Test Student',
  email: '',
  grade,
  subjects: ['数'],
  subjectSlots: { 数: 1 },
  unavailableDates: [],
  preferredSlots: [],
  unavailableSlots: [],
  memo: '',
  submittedAt: 1,
  constraintCards,
})

describe('evaluateConstraintCards slot preferences', () => {
  it('prefers elementary slot 2 over 3 over 4 over 5, and avoids slot 1', () => {
    const student = makeStudent('小4', ['earlySlotPreference'])
    const assignments = {}

    const score1 = evaluateConstraintCards(student, '2026-03-23_1', assignments, 6, [], []).score
    const score2 = evaluateConstraintCards(student, '2026-03-23_2', assignments, 6, [], []).score
    const score3 = evaluateConstraintCards(student, '2026-03-23_3', assignments, 6, [], []).score
    const score4 = evaluateConstraintCards(student, '2026-03-23_4', assignments, 6, [], []).score
    const score5 = evaluateConstraintCards(student, '2026-03-23_5', assignments, 6, [], []).score

    expect(score2).toBeGreaterThan(score3)
    expect(score3).toBeGreaterThan(score4)
    expect(score4).toBeGreaterThan(score5)
    expect(score5).toBeGreaterThan(score1)
  })

  it('strongly prefers secondary slot 5 over 4 and penalizes early slots', () => {
    const student = makeStudent('中2', ['lateSlotPreference'])
    const assignments = {}

    const score1 = evaluateConstraintCards(student, '2026-03-23_1', assignments, 6, [], []).score
    const score2 = evaluateConstraintCards(student, '2026-03-23_2', assignments, 6, [], []).score
    const score3 = evaluateConstraintCards(student, '2026-03-23_3', assignments, 6, [], []).score
    const score4 = evaluateConstraintCards(student, '2026-03-23_4', assignments, 6, [], []).score
    const score5 = evaluateConstraintCards(student, '2026-03-23_5', assignments, 6, [], []).score

    expect(score5).toBeGreaterThan(score4)
    expect(score4).toBeGreaterThan(score3)
    expect(score3).toBeGreaterThan(score2)
    expect(score2).toBeGreaterThan(score1)
  })
})