import { Component, type ReactNode } from 'react'
import { WindowEmpty } from '../../ui/WindowEmpty'
import { readySlides, undoSlides } from './slides-store'

interface Props {
  readonly documentId: string
  readonly children: ReactNode
}

interface State {
  readonly failed: boolean
}

/**
 * Keeps a slide that cannot be drawn from blanking the whole window: the tab says so instead, and
 * offers to undo the change that caused it. React only catches drawing errors in a class component.
 */
export class SlidesErrorBoundary extends Component<Props, State> {
  override state: State = { failed: false }

  static getDerivedStateFromError(): State {
    return { failed: true }
  }

  override render(): ReactNode {
    if (!this.state.failed) return this.props.children
    const { documentId } = this.props
    const canUndo = (readySlides(documentId)?.past.length ?? 0) > 0
    return (
      <WindowEmpty
        application="slides"
        title="This presentation could not be shown"
        description={canUndo ? 'The last change left something Slides cannot draw. Undo it to carry on.' : 'Something in it cannot be drawn. Close the tab; the file itself is unchanged.'}
      >
        {canUndo && (
          <button type="button" className="button" onClick={() => {
            undoSlides(documentId)
            this.setState({ failed: false })
          }}>
            Undo last change
          </button>
        )}
      </WindowEmpty>
    )
  }
}
