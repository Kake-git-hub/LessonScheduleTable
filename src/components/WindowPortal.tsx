import { createPortal } from 'react-dom'
import { useEffect, useRef, useState, type ReactNode } from 'react'

type WindowPortalProps = {
  children: ReactNode
  onClose: () => void
  title?: string
  width?: number
  height?: number
}

export default function WindowPortal({
  children,
  onClose,
  title = 'コマ調整',
  width = 1400,
  height = 800,
}: WindowPortalProps) {
  const [container, setContainer] = useState<HTMLElement | null>(null)
  const externalWindowRef = useRef<Window | null>(null)
  const onCloseRef = useRef(onClose)
  onCloseRef.current = onClose

  useEffect(() => {
    const externalWindow = window.open('', 'slot-adjust-window')
    if (!externalWindow) {
      alert('ポップアップがブロックされました。ポップアップを許可してください。')
      onCloseRef.current()
      return
    }

    externalWindowRef.current = externalWindow
    externalWindow.document.title = title

    // Copy styles from parent window
    for (const sheet of Array.from(document.styleSheets)) {
      try {
        const rules = Array.from(sheet.cssRules).map((r) => r.cssText).join('\n')
        const style = externalWindow.document.createElement('style')
        style.textContent = rules
        externalWindow.document.head.appendChild(style)
      } catch {
        if (sheet.href) {
          const link = externalWindow.document.createElement('link')
          link.rel = 'stylesheet'
          link.href = sheet.href
          externalWindow.document.head.appendChild(link)
        }
      }
    }

    // Mount container
    const div = externalWindow.document.createElement('div')
    div.id = 'slot-adjust-root'
    externalWindow.document.body.style.margin = '0'
    externalWindow.document.body.appendChild(div)
    setContainer(div)

    // Detect external window close
    const interval = setInterval(() => {
      if (externalWindow.closed) {
        clearInterval(interval)
        onCloseRef.current()
      }
    }, 300)

    return () => {
      clearInterval(interval)
      if (!externalWindow.closed) externalWindow.close()
    }
    // mount-only
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  if (!container) return null
  return createPortal(children, container)
}
