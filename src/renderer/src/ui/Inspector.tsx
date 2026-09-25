import type { ReactNode } from 'react'

export function Inspector({ label, children }: { readonly label: string; readonly children: ReactNode }) {
  return <aside className="inspector" aria-label={label}>{children}</aside>
}

export function InspectorSection({ title, children }: { readonly title?: string; readonly children: ReactNode }) {
  return (
    <section className="inspector-section">
      {title !== undefined && <h3>{title}</h3>}
      {children}
    </section>
  )
}

export function InspectorRow({ label, children }: { readonly label: string; readonly children: ReactNode }) {
  return <div className="inspector-row"><span>{label}</span>{children}</div>
}
