import { ContextMenu } from '@base-ui/react/context-menu'
import { closestCenter, DndContext, KeyboardSensor, PointerSensor, useSensor, useSensors, type DragEndEvent } from '@dnd-kit/core'
import { SortableContext, sortableKeyboardCoordinates, useSortable, verticalListSortingStrategy } from '@dnd-kit/sortable'
import { useEffect, useRef } from 'react'
import { commandShortcut } from '../../ui/Tip'
import type { Slide } from './model'
import { addSlide, deleteSlide, duplicateCurrentSlide, moveSlide, showSlide } from './slides-actions'
import type { ReadySlides } from './slides-store'
import { SlideView } from './SlideView'

const thumbnailWidth = 148

interface ThumbnailProps {
  readonly documentId: string
  readonly document: ReadySlides
  readonly slide: Slide
  readonly position: number
}

function Thumbnail({ documentId, document, slide, position }: ThumbnailProps) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: slide.id })
  const current = slide.id === document.slideId
  const { width, height } = document.present
  const ref = useRef<HTMLElement | null>(null)

  // Keeps the current slide in view as it changes from the keyboard, a new slide or undo.
  useEffect(() => {
    if (current) ref.current?.scrollIntoView({ block: 'nearest' })
  }, [current])

  return (
    <ContextMenu.Root>
      <ContextMenu.Trigger
        ref={element => {
          setNodeRef(element)
          ref.current = element
        }}
        className={`slides-thumb${current ? ' is-current' : ''}${isDragging ? ' is-dragging' : ''}`}
        style={transform === null ? undefined : { transform: `translate3d(0, ${transform.y}px, 0)`, transition }}
        {...attributes}
        {...listeners}
        aria-label={`Slide ${position + 1}`}
        aria-current={current ? 'page' : undefined}
        onClick={() => showSlide(documentId, slide.id)}
        onContextMenu={() => showSlide(documentId, slide.id)}
      >
        <span className="slides-thumb-number">{position + 1}</span>
        <span className="slides-thumb-page">
          <SlideView slide={slide} width={width} height={height} scale={thumbnailWidth / width} />
        </span>
      </ContextMenu.Trigger>
      <ContextMenu.Portal>
        <ContextMenu.Positioner>
          <ContextMenu.Popup className="menu-popup">
            <ContextMenu.Item className="menu-item" onClick={() => addSlide(documentId, 'titleContent')}>
              <span>New Slide</span><kbd>{commandShortcut('slide.new')}</kbd>
            </ContextMenu.Item>
            <ContextMenu.Item className="menu-item" onClick={() => duplicateCurrentSlide(documentId, slide.id)}><span>Duplicate Slide</span></ContextMenu.Item>
            <ContextMenu.Separator className="menu-separator" />
            <ContextMenu.Item className="menu-item" disabled={document.present.slides.length === 1} onClick={() => deleteSlide(documentId, slide.id)}>
              <span>Delete Slide</span>
            </ContextMenu.Item>
          </ContextMenu.Popup>
        </ContextMenu.Positioner>
      </ContextMenu.Portal>
    </ContextMenu.Root>
  )
}

/** Slide thumbnails: click to show a slide, drag to reorder, right-click for slide actions. */
export function SlideRail({ documentId, document }: { readonly documentId: string; readonly document: ReadySlides }) {
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 5 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates }),
  )
  const handleDragEnd = (event: DragEndEvent): void => {
    if (event.over !== null) moveSlide(documentId, String(event.active.id), String(event.over.id))
  }

  return (
    <nav className="slides-rail" aria-label="Slides">
      <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
        <SortableContext items={document.present.slides.map(slide => slide.id)} strategy={verticalListSortingStrategy}>
          {document.present.slides.map((slide, position) => (
            <Thumbnail key={slide.id} documentId={documentId} document={document} slide={slide} position={position} />
          ))}
        </SortableContext>
      </DndContext>
    </nav>
  )
}
