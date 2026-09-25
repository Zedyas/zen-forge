import { useState } from 'react'
import { closestCenter, DndContext, KeyboardSensor, PointerSensor, useSensor, useSensors, type DragEndEvent } from '@dnd-kit/core'
import { SortableContext, sortableKeyboardCoordinates, useSortable, verticalListSortingStrategy } from '@dnd-kit/sortable'
import type { PageItem } from './model'
import { PageCanvas } from './PageCanvas'
import { movePage } from './pdf-actions'
import { scrollToPage } from './PdfDesk'
import { pageSize, updatePdf, type ReadyPdf } from './pdf-store'

const thumbnailWidth = 88

interface ThumbnailProps {
  readonly documentId: string
  readonly document: ReadyPdf
  readonly item: PageItem
  readonly position: number
  readonly root: Element | null
}

function Thumbnail({ documentId, document, item, position, root }: ThumbnailProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: item.key })
  const size = pageSize(document, item)
  const scale = thumbnailWidth / size.width
  const source = document.sources[item.source]
  const current = position === document.currentPage

  return (
    <div
      ref={setNodeRef}
      className={`pdf-thumb${current ? ' is-current' : ''}${isDragging ? ' is-dragging' : ''}`}
      style={transform === null ? undefined : { transform: `translate3d(0, ${transform.y}px, 0)`, transition }}
      {...attributes}
      {...listeners}
      aria-label={`Page ${position + 1}`}
      aria-current={current ? 'page' : undefined}
      onClick={() => {
        updatePdf(documentId, () => ({ currentPage: position }))
        scrollToPage(item.key)
      }}
    >
      <div className="pdf-thumb-page" style={{ width: thumbnailWidth, height: size.height * scale }}>
        {source !== undefined && (
          <PageCanvas source={source} index={item.index} extraRotation={item.rotation} scale={scale} forms="print" root={root} margin={400} />
        )}
      </div>
      <span className="pdf-thumb-label">{position + 1}</span>
    </div>
  )
}

/** Page thumbnails: click to go to a page, drag to reorder. */
export function PageRail({ documentId, document }: { readonly documentId: string; readonly document: ReadyPdf }) {
  const [rail, setRail] = useState<HTMLElement | null>(null)
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates }),
  )
  const handleDragEnd = (event: DragEndEvent): void => {
    if (event.over !== null) movePage(documentId, String(event.active.id), String(event.over.id))
  }

  return (
    <nav ref={setRail} className="pdf-rail" aria-label="Pages">
      <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
        <SortableContext items={document.present.pages.map(item => item.key)} strategy={verticalListSortingStrategy}>
          {document.present.pages.map((item, position) => (
            <Thumbnail key={item.key} documentId={documentId} document={document} item={item} position={position} root={rail} />
          ))}
        </SortableContext>
      </DndContext>
    </nav>
  )
}
