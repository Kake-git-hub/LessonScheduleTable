/**
 * 3-level subject system for teachers.
 *
 * Teachers register subjects with a level prefix: 小英, 中英, 高英, etc.
 * The hierarchy is: 高 > 中 > 小.
 *   - 高英 means: can teach 英 to all grades (小/中/高)
 *   - 中英 means: can teach 英 to 小 and 中 grades only
 *   - 小英 means: can teach 英 to 小 grades only
 *
 * Students continue to use base subjects: 英, 数, 国, etc.
 * Matching uses the student's grade to determine the required level.
 *
 * Legacy subjects without prefix (e.g. '英') are treated as 高-level for backward compatibility.
 */

/** Base (non-leveled) subjects used by students. */
export const BASE_SUBJECTS = ['英', '数', '国', '理', '社', 'IT', '算'] as const

/** Combo subjects for elementary students (two subjects in one slot). */
export const ELEMENTARY_COMBO_SUBJECTS = ['算英', '算国', '英国'] as const

/** Level prefixes ordered from lowest to highest. */
export const LEVEL_PREFIXES = ['小', '中', '高'] as const

/** All leveled teacher subject options (小英, 中英, 高英, ...). */
export const TEACHER_SUBJECTS: string[] = LEVEL_PREFIXES.flatMap(lv =>
  BASE_SUBJECTS.map(s => `${lv}${s}`),
)

const LEVEL_ORDER: Record<string, number> = { '小': 0, '中': 1, '高': 2 }

/** Extract the base subject from a (possibly leveled) subject string.
 *  '高英' → '英', '中数' → '数', '英' → '英' (legacy). */
export const getSubjectBase = (subj: string): string => {
  if (LEVEL_PREFIXES.some(lv => subj.startsWith(lv))) {
    return subj.slice(1)
  }
  return subj // legacy: no prefix
}

/** Extract the level prefix from a (possibly leveled) subject string.
 *  '高英' → '高', '中数' → '中', '英' → '高' (legacy treated as highest). */
export const getSubjectLevel = (subj: string): string => {
  for (const lv of LEVEL_PREFIXES) {
    if (subj.startsWith(lv)) return lv
  }
  return '高' // legacy: treat as highest level
}

/** Map a student grade string to a level prefix.
 *  '高1' → '高', '中2' → '中', '小3' → '小'. */
export const gradeToLevel = (grade: string): string => {
  if (grade.startsWith('高')) return '高'
  if (grade.startsWith('中')) return '中'
  return '小'
}

/**
 * Check if a teacher (given their subject list) can teach a base subject
 * to a student of the given grade.
 * For combo subjects (算英, 算国, 英国), the teacher must be able to teach
 * both component subjects at the student's grade level.
 *
 * Example: canTeachSubject(['高英', '中数'], '中2', '英') → true
 *          canTeachSubject(['中数'], '高1', '数') → false
 *          canTeachSubject(['小算', '小英'], '小3', '算英') → true
 */
export const canTeachSubject = (
  teacherSubjects: string[],
  studentGrade: string,
  baseSubject: string,
): boolean => {
  // Handle combo subjects
  if ((ELEMENTARY_COMBO_SUBJECTS as readonly string[]).includes(baseSubject)) {
    const components = [...baseSubject] // e.g. '算英' → ['算', '英']
    return components.every(comp => canTeachSubject(teacherSubjects, studentGrade, comp))
  }
  const requiredLevel = LEVEL_ORDER[gradeToLevel(studentGrade)] ?? 0
  return teacherSubjects.some(ts => {
    const base = getSubjectBase(ts)
    if (base !== baseSubject) return false
    const teacherLevel = LEVEL_ORDER[getSubjectLevel(ts)] ?? 0
    return teacherLevel >= requiredLevel
  })
}

/**
 * Get all base subjects a teacher can teach to a student of the given grade.
 * For elementary students, also includes available combo subjects (算英, 算国, 英国).
 *
 * Returns an array of subject names (e.g. ['英', '数', '算', '算英']).
 */
export const teachableBaseSubjects = (
  teacherSubjects: string[],
  studentGrade: string,
): string[] => {
  const singles = (BASE_SUBJECTS as readonly string[]).filter(base =>
    canTeachSubject(teacherSubjects, studentGrade, base),
  )
  // Add combo subjects for elementary students
  if (gradeToLevel(studentGrade) === '小') {
    for (const combo of ELEMENTARY_COMBO_SUBJECTS) {
      if (canTeachSubject(teacherSubjects, studentGrade, combo)) {
        singles.push(combo)
      }
    }
  }
  return singles
}

/**
 * Check if a teacher has ANY variant of a base subject (regardless of level).
 * For combo subjects, checks that the teacher has both component subjects.
 * Useful when grade context is not available.
 *
 * Example: teacherHasSubject(['高英', '中数'], '英') → true
 *          teacherHasSubject(['小算', '小英'], '算英') → true
 */
export const teacherHasSubject = (
  teacherSubjects: string[],
  baseSubject: string,
): boolean => {
  if ((ELEMENTARY_COMBO_SUBJECTS as readonly string[]).includes(baseSubject)) {
    return [...baseSubject].every(comp => teacherHasSubject(teacherSubjects, comp))
  }
  return teacherSubjects.some(ts => getSubjectBase(ts) === baseSubject)
}
