import { ContextMenu } from '@base-ui/react/context-menu'
import { useSortable } from '@dnd-kit/sortable'
import { X } from 'lucide-react'
import { toast } from 'sonner'
import { closeDocument, moveToNewWindow } from '../app/document-actions'
import { useDocumentsStore, type OpenDocument } from '../app/documents-store'
import { platformClient } from '../services/platform/client'
import { AppIcon } from './AppIcon'

function reportFailure(error: unknown): void {
  toast.error('Could not move the tab', { description: error instanceof Error ? error.message : undefined })
}

/** A document tab: its application's icon says what kind of file it is; right-click for window actions. */
export function DocumentTab({ document, active }: { readonly document: OpenDocument; readonly active: boolean }) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: document.id })
  const style = transform === null ? undefined : { transform: `translate3d(${transform.x}px, 0, 0)`, transition }

  return (
    <ContextMenu.Root>
      <ContextMenu.Trigger
        ref={setNodeRef}
        style={style}
        className={`doc-tab${active ? ' is-active' : ''}${isDragging ? ' is-dragging' : ''}`}
        {...attributes}
        {...listeners}
        role="tab"
        aria-selected={active}
        title={document.path ?? document.name}
        onClick={() => useDocumentsStore.getState().select(document.id)}
        onAuxClick={event => {
          if (event.button === 1) void closeDocument(document.id)
        }}
      >
        <AppIcon application={document.kind} size={15} />
        <span className="doc-tab-name">{document.name}</span>
        <button
          type="button"
          className={`doc-tab-close${document.dirty ? ' is-dirty' : ''}`}
          aria-label={document.dirty ? `Close ${document.name} (unsaved changes)` : `Close ${document.name}`}
          onPointerDown={event => event.stopPropagation()}
          onClick={event => {
            event.stopPropagation()
            void closeDocument(document.id)
          }}
        >
          {document.dirty && <i className="dirty-dot" aria-hidden="true" />}
          <X aria-hidden="true" size={12} strokeWidth={2} />
        </button>
      </ContextMenu.Trigger>
      <ContextMenu.Portal>
        <ContextMenu.Positioner>
          <ContextMenu.Popup className="menu-popup">
            <ContextMenu.Item className="menu-item" onClick={() => void moveToNewWindow(document.id).catch(reportFailure)}>
              <span>Move to New Window</span>
            </ContextMenu.Item>
            {document.path !== undefined && (
              <ContextMenu.Item className="menu-item" onClick={() => void platformClient.revealInFinder(document.path ?? '')}>
                <span>Show in Finder</span>
              </ContextMenu.Item>
            )}
            <ContextMenu.Separator className="menu-separator" />
            <ContextMenu.Item className="menu-item" onClick={() => void closeDocument(document.id)}>
              <span>Close Tab</span><kbd>⌘W</kbd>
            </ContextMenu.Item>
          </ContextMenu.Popup>
        </ContextMenu.Positioner>
      </ContextMenu.Portal>
    </ContextMenu.Root>
  )
}
