import { createContext, useContext, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { useViewStore } from '../app/view-store'

/** The title-bar element that editors fill with their essential tools while the toolbar is hidden. */
export const TitleSlotContext = createContext<HTMLElement | null>(null)

/** Renders the minimal-mode tool group in the title bar; nothing while the full toolbar is shown. */
export function TitleEssentials({ children }: { readonly children: ReactNode }) {
  const slot = useContext(TitleSlotContext)
  const toolbar = useViewStore(state => state.toolbar)
  if (slot === null || toolbar) return null
  return createPortal(<div className="title-essentials no-drag">{children}</div>, slot)
}
