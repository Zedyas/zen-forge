import type { ReactNode } from 'react'
import type { EditorApplicationId } from '@shared/applications'
import { AppIcon } from './AppIcon'

interface WindowEmptyProps {
  readonly application: EditorApplicationId
  readonly title: string
  readonly description: string
  readonly children?: ReactNode
}

/** A tab whose document could not be shown: say what happened and offer the next step. */
export function WindowEmpty({ application, title, description, children }: WindowEmptyProps) {
  return (
    <section className="window-empty">
      <AppIcon application={application} size={56} />
      <h2>{title}</h2>
      <p>{description}</p>
      {children !== undefined && <div className="actions">{children}</div>}
    </section>
  )
}
