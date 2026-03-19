const JAPANESE_FONT_STYLE_ID = 'lesson-schedule-japanese-font'
const JAPANESE_FONT_FILE_NAME = 'BIZUDPGothic-Regular.ttf'
const JAPANESE_FONT_INTERNAL_FAMILY = 'LessonScheduleJapanese'

export const JAPANESE_FONT_FAMILY = `"${JAPANESE_FONT_INTERNAL_FAMILY}", "BIZ UDPGothic", "BIZ UDP Gothic", "Yu Gothic", "YuGothic", "Hiragino Sans", "Meiryo", sans-serif`

function normalizeBaseUrl(baseUrl?: string): string {
  const resolvedBaseUrl = baseUrl ?? import.meta.env.BASE_URL ?? '/'
  return resolvedBaseUrl.endsWith('/') ? resolvedBaseUrl : `${resolvedBaseUrl}/`
}

export function getJapaneseFontUrl(baseUrl?: string): string {
  return `${normalizeBaseUrl(baseUrl)}fonts/${JAPANESE_FONT_FILE_NAME}`
}

export function getJapaneseFontFaceCss(baseUrl?: string): string {
  return `@font-face {
  font-family: "${JAPANESE_FONT_INTERNAL_FAMILY}";
  src: url("${getJapaneseFontUrl(baseUrl)}") format("truetype");
  font-style: normal;
  font-weight: 400;
  font-display: swap;
}`
}

export function ensureJapaneseFontStyle(doc: Document = document, baseUrl?: string): void {
  if (doc.getElementById(JAPANESE_FONT_STYLE_ID)) return

  const style = doc.createElement('style')
  style.id = JAPANESE_FONT_STYLE_ID
  style.textContent = getJapaneseFontFaceCss(baseUrl)
  ;(doc.head ?? doc.documentElement).appendChild(style)
}

export async function waitForJapaneseFontReady(doc: Document = document, baseUrl?: string): Promise<void> {
  ensureJapaneseFontStyle(doc, baseUrl)

  if (!('fonts' in doc) || !doc.fonts) return

  try {
    await doc.fonts.load(`400 1em ${JAPANESE_FONT_INTERNAL_FAMILY}`)
    await doc.fonts.ready
  } catch {
    // Ignore font loading errors and allow existing fallbacks to render.
  }
}
