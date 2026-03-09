/**
 * Teacher subject system.
 *
 * Teachers register subjects with a level prefix: 小英, 中英, 高1英, 高2英, 高3英, etc.
 * The hierarchy is: 高3 > 高2 > 高1 > 中 > 小.
 *   - 高3英 means: can teach 英 to all grades up to 高3
 *   - 高2英 means: can teach 英 to all grades up to 高2
 *   - 高1英 means: can teach 英 to all grades up to 高1
 *   - 中英 means: can teach 英 to 小 and 中 grades only
 *   - 小英 means: can teach 英 to 小 grades only
 *
 * Students continue to use base subjects: 英, 数, 国, etc.
 * Matching uses the student's grade to determine the required level.
 *
 * Backward compatibility:
 *   - Legacy 高英 / 高数 / ... are treated as the broadest high-school tier.
 *   - Legacy non-prefixed subjects (e.g. '英') are treated the same as legacy 高英.
 */

/** Base (non-leveled) subjects used by students. */
export const BASE_SUBJECTS = ['英', '数', '国', '理', '社', 'IT', '算'] as const

/** Combo subjects for elementary students (two subjects in one slot). */
export const ELEMENTARY_COMBO_SUBJECTS = ['算英', '算国', '英国'] as const

/** Selectable level prefixes ordered from lowest to highest. */
export const LEVEL_PREFIXES = ['小', '中', '高1', '高2', '高3'] as const

/** Legacy prefixes kept only for backward compatibility with existing data/imports. */
const LEGACY_LEVEL_PREFIXES = ['高'] as const

const RECOGNIZED_LEVEL_PREFIXES = [...LEVEL_PREFIXES, ...LEGACY_LEVEL_PREFIXES] as const

/** 算 and 数 are equivalent (算数 = elementary math, 数学 = secondary math). */
const EQUIVALENT_SUBJECTS: Record<string, string> = { '算': '数', '数': '算' }

/** Ordered base subjects per level for TEACHER_SUBJECTS generation.
 *  小: 英, 算, 国, 理, 社, IT  (算 replaces 数)
 *  中/高1/高2/高3: 英, 数, 国, 理, 社, IT */
const subjectsForLevel = (lv: string): string[] => {
  if (lv === '小') return ['英', '算', '国', '理', '社', 'IT']
  return ['英', '数', '国', '理', '社', 'IT']
}

/** All selectable teacher subject options (小英, 中英, 高1英, 高2英, 高3英, ...). */
export const TEACHER_SUBJECTS: string[] = LEVEL_PREFIXES.flatMap(lv =>
  subjectsForLevel(lv).map(s => `${lv}${s}`),
)

const LEGACY_TEACHER_SUBJECTS: string[] = LEGACY_LEVEL_PREFIXES.flatMap(lv =>
  subjectsForLevel(lv).map(s => `${lv}${s}`),
)

const KNOWN_TEACHER_SUBJECTS = new Set<string>([
  ...TEACHER_SUBJECTS,
  ...LEGACY_TEACHER_SUBJECTS,
])

const LEVEL_ORDER: Record<string, number> = { '小': 0, '中': 1, '高1': 2, '高2': 3, '高3': 4, '高': 4 }

const SORTED_LEVEL_PREFIXES = [...RECOGNIZED_LEVEL_PREFIXES].sort((a, b) => b.length - a.length)

const parseTeacherSubject = (subj: string): { level: string; base: string } | null => {
  for (const lv of SORTED_LEVEL_PREFIXES) {
    if (subj.startsWith(lv)) {
      return { level: lv, base: subj.slice(lv.length) }
    }
  }
  return null
}

export const isKnownTeacherSubject = (subj: string): boolean => KNOWN_TEACHER_SUBJECTS.has(subj)

/** Extract the base subject from a (possibly leveled) subject string.
 *  '高2英' → '英', '高英' → '英', '中数' → '数', '英' → '英' (legacy). */
export const getSubjectBase = (subj: string): string => {
  const parsed = parseTeacherSubject(subj)
  if (parsed) {
    return parsed.base
  }
  return subj // legacy: no prefix
}

/** Extract the level prefix from a (possibly leveled) subject string.
 *  '高2英' → '高2', '高英' → '高', '中数' → '中', '英' → '高' (legacy treated as highest). */
export const getSubjectLevel = (subj: string): string => {
  const parsed = parseTeacherSubject(subj)
  if (parsed) {
    return parsed.level
  }
  return '高' // legacy: treat as highest level
}

/** Map a student grade string to a level prefix.
 *  '高1' → '高1', '高2' → '高2', '高3' → '高3', '中2' → '中', '小3' → '小'. */
export const gradeToLevel = (grade: string): string => {
  if (grade.startsWith('高3')) return '高3'
  if (grade.startsWith('高2')) return '高2'
  if (grade.startsWith('高1')) return '高1'
  if (grade.startsWith('高')) return '高3'
  if (grade.startsWith('中')) return '中'
  return '小'
}

/**
 * Check if a teacher (given their subject list) can teach a base subject
 * to a student of the given grade.
 * For combo subjects (算英, 算国, 英国), the teacher must be able to teach
 * both component subjects at the student's grade level.
 *
 * Example: canTeachSubject(['高2英', '中数'], '中2', '英') → true
 *          canTeachSubject(['高2数'], '高3', '数') → false
 *          canTeachSubject(['高2数'], '高1', '数') → true
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
  const equiv = EQUIVALENT_SUBJECTS[baseSubject]
  return teacherSubjects.some(ts => {
    const base = getSubjectBase(ts)
    if (base !== baseSubject && base !== equiv) return false
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
 * Example: teacherHasSubject(['高2英', '中数'], '英') → true
 *          teacherHasSubject(['小算', '小英'], '算英') → true
 */
export const teacherHasSubject = (
  teacherSubjects: string[],
  baseSubject: string,
): boolean => {
  if ((ELEMENTARY_COMBO_SUBJECTS as readonly string[]).includes(baseSubject)) {
    return [...baseSubject].every(comp => teacherHasSubject(teacherSubjects, comp))
  }
  const equiv = EQUIVALENT_SUBJECTS[baseSubject]
  return teacherSubjects.some(ts => {
    const base = getSubjectBase(ts)
    return base === baseSubject || base === equiv
  })
}
