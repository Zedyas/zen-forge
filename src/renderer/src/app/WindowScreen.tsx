import { useEffect, useRef, useState, type DragEvent } from 'react'
import { Menu } from '@base-ui/react/menu'
import { closestCenter, DndContext, PointerSensor, useSensor, useSensors, type DragEndEvent } from '@dnd-kit/core'
import { horizontalListSortingStrategy, SortableContext } from '@dnd-kit/sortable'
import { FilePlus2, Files, FolderOpen, Images, PanelRight, PanelTop, Plus } from 'lucide-react'
import { toast } from 'sonner'
import { suiteName } from '@shared/applications'
import type { CommandId } from '@shared/commands'
import { fileService } from '../services/file/IpcFileService'
import { recordRecent } from '../services/index/document-index'
import { platformClient } from '../services/platform/client'
import { AppIcon } from '../ui/AppIcon'
import { DocumentTab } from '../ui/DocumentTab'
import { icon } from '../ui/icons'
import { commandShortcut, Tip } from '../ui/Tip'
import { TitleSlotContext } from '../ui/TitleSlot'
import { DocsWorkspace } from '../modules/docs/DocsWorkspace'
import { PdfWorkspace } from '../modules/pdf/PdfWorkspace'
import { combinePdfs, createPdfFromImages } from '../modules/pdf/workflows'
import { SheetsWorkspace } from '../modules/sheets/SheetsWorkspace'
import { SlidesWorkspace } from '../modules/slides/SlidesWorkspace'
import { executeCommand } from './commands'
import { openDroppedFiles, resolveWindowClose } from './document-actions'
import { activeDocument, homeTabId, useDocumentsStore } from './documents-store'
import { HomeScreen } from './HomeScreen'
import { useViewStore } from './view-store'

function reportFailure(title: string) {
  return (error: unknown): void => {
    toast.error(title, { description: error instanceof Error ? error.message : undefined })
  }
}

function runCommand(command: CommandId): void {
  void executeCommand(command).catch(reportFailure('Command failed'))
}

/** Restores last session's tabs, then keeps the main process told about title, menus, unsaved state and session. */
function useWindowSession(): void {
  // Set when a file arrives from Finder or a drop while last session is still being restored.
  const openedExplicitly = useRef(false)

  useEffect(() => {
    void (async () => {
      const session = await platformClient.takeSession()
      if (session === undefined) return
      for (const path of session.paths) {
        try {
          useDocumentsStore.getState().restoreFile(await fileService.describe(path))
        } catch {
          // A file moved or deleted since last session is skipped; recents still list where it was.
        }
      }
      // An explicitly opened file keeps focus over last session's active tab.
      if (!openedExplicitly.current && session.activePath !== undefined) {
        useDocumentsStore.getState().select(session.activePath)
      }
    })()
  }, [])

  useEffect(() => platformClient.onDocumentOpened(file => {
    openedExplicitly.current = true
    useDocumentsStore.getState().openFile(file)
    void recordRecent(file)
  }), [])

  useEffect(() => platformClient.onCloseRequested(() => void resolveWindowClose()), [])

  useEffect(() => {
    const sync = (): void => {
      const state = useDocumentsStore.getState()
      const active = activeDocument(state)
      // Portalled menus and dialogs read the accent from <html>, so it follows the active tab too.
      document.documentElement.dataset['app'] = active?.kind ?? 'home'
      void platformClient.setWindowState({
        title: active?.name ?? suiteName,
        active: active?.kind ?? 'home',
        path: active?.path,
        edited: active?.dirty ?? false,
        hasUnsaved: state.documents.some(candidate => candidate.dirty),
        session: {
          paths: state.documents.flatMap(candidate => candidate.path === undefined ? [] : [candidate.path]),
          activePath: active?.path,
        },
      })
    }
    sync()
    return useDocumentsStore.subscribe(sync)
  }, [])

  useEffect(() => {
    const { toolbar, inspector } = useViewStore.getState()
    void platformClient.setViewState({ toolbar, inspector })
  }, [])
}

function HomeTab({ active }: { readonly active: boolean }) {
  return (
    <Tip label="Home" shortcut={commandShortcut('shell.home')}>
      <button
        type="button"
        role="tab"
        aria-selected={active}
        aria-label="Home"
        className={`home-tab${active ? ' is-active' : ''}`}
        onClick={() => useDocumentsStore.getState().select(homeTabId)}
      >
        {/* A size up from the tab icons, so the suite mark reads as the window's own tab. */}
        <AppIcon application="home" size={20} />
      </button>
    </Tip>
  )
}

function NewTabMenu() {
  return (
    <Menu.Root>
      <Tip label="New tab">
        <Menu.Trigger className="tool new-tab" aria-label="New tab">
          <Plus {...icon} />
        </Menu.Trigger>
      </Tip>
      <Menu.Portal>
        <Menu.Positioner sideOffset={6} align="start">
          <Menu.Popup className="menu-popup">
            <Menu.Item className="menu-item" onClick={() => runCommand('file.new')}>
              <AppIcon application="sheets" size={icon.size} /><span>New spreadsheet</span><kbd>{commandShortcut('file.new')}</kbd>
            </Menu.Item>
            <Menu.Item className="menu-item" onClick={() => runCommand('file.newDocument')}>
              <AppIcon application="docs" size={20} /><span>New document</span>
            </Menu.Item>
            <Menu.Item className="menu-item" onClick={() => runCommand('file.newPresentation')}>
              <AppIcon application="slides" size={20} /><span>New presentation</span>
            </Menu.Item>
            <Menu.Item className="menu-item" onClick={() => runCommand('file.open')}>
              <FolderOpen {...icon} /><span>Open…</span><kbd>{commandShortcut('file.open')}</kbd>
            </Menu.Item>
            <Menu.Separator className="menu-separator" />
            <Menu.Item className="menu-item" onClick={() => void createPdfFromImages().catch(reportFailure('Could not create the PDF'))}>
              <Images {...icon} /><span>PDF from images…</span>
            </Menu.Item>
            <Menu.Item className="menu-item" onClick={() => void combinePdfs().catch(reportFailure('Could not combine the PDFs'))}>
              <Files {...icon} /><span>Combine PDFs…</span>
            </Menu.Item>
          </Menu.Popup>
        </Menu.Positioner>
      </Menu.Portal>
    </Menu.Root>
  )
}

function ViewToggles() {
  const toolbar = useViewStore(state => state.toolbar)
  const inspector = useViewStore(state => state.inspector)
  return (
    <div className="title-actions">
      <Tip label={toolbar ? 'Hide toolbar' : 'Show toolbar'} shortcut={commandShortcut('view.toolbar')}>
        <button type="button" className="tool" aria-label="Toolbar" aria-pressed={toolbar} onClick={() => runCommand('view.toolbar')}>
          <PanelTop {...icon} />
        </button>
      </Tip>
      <Tip label={inspector ? 'Hide inspector' : 'Show inspector'} shortcut={commandShortcut('view.inspector')}>
        <button type="button" className="tool" aria-label="Inspector" aria-pressed={inspector} onClick={() => runCommand('view.inspector')}>
          <PanelRight {...icon} />
        </button>
      </Tip>
    </div>
  )
}

/** Only file drags from Finder count; text drags inside the page do not. */
function isFileDrag(event: DragEvent): boolean {
  return event.dataTransfer.types.includes('Files')
}

/** One window: the pinned Home tab, document tabs of any kind, and the active tab's editor below. */
export function WindowScreen() {
  const documents = useDocumentsStore(state => state.documents)
  const activeId = useDocumentsStore(state => state.activeId)
  const toolbar = useViewStore(state => state.toolbar)
  const [titleSlot, setTitleSlot] = useState<HTMLElement | null>(null)
  const [dragging, setDragging] = useState(false)
  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 5 } }))
  useWindowSession()

  const active = documents.find(document => document.id === activeId)

  const handleDragEnd = (event: DragEndEvent): void => {
    if (event.over !== null) useDocumentsStore.getState().reorder(String(event.active.id), String(event.over.id))
  }

  return (
    <TitleSlotContext.Provider value={titleSlot}>
      <main
        className="window"
        data-toolbar={toolbar || active === undefined ? 'shown' : 'hidden'}
        onDragEnter={event => {
          if (!isFileDrag(event)) return
          event.preventDefault()
          setDragging(true)
        }}
        onDragOver={event => {
          if (isFileDrag(event)) event.preventDefault()
        }}
        onDragLeave={event => {
          if (event.relatedTarget === null) setDragging(false)
        }}
        onDrop={event => {
          if (!isFileDrag(event)) return
          event.preventDefault()
          setDragging(false)
          void openDroppedFiles(event.dataTransfer.files).catch(reportFailure('Could not open the dropped file'))
        }}
      >
        <header className="titlebar titlebar-drag">
          <div className="titlebar-tabs">
            <HomeTab active={active === undefined} />
            <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEnd}>
              <div className="doc-tabs" role="tablist" aria-label="Open files">
                <SortableContext items={documents.map(document => document.id)} strategy={horizontalListSortingStrategy}>
                  {documents.map(document => <DocumentTab key={document.id} document={document} active={document.id === activeId} />)}
                </SortableContext>
              </div>
            </DndContext>
            <NewTabMenu />
          </div>
          <div ref={setTitleSlot} className="no-drag" />
          {active !== undefined && <ViewToggles />}
        </header>

        {active === undefined && <HomeScreen />}
        {active?.kind === 'sheets' && <SheetsWorkspace document={active} />}
        {active?.kind === 'pdf' && <PdfWorkspace document={active} />}
        {active?.kind === 'docs' && <DocsWorkspace document={active} />}
        {active?.kind === 'slides' && <SlidesWorkspace document={active} />}

        {dragging && <div className="drop-overlay"><FilePlus2 size={28} aria-hidden="true" />Drop to open</div>}
      </main>
    </TitleSlotContext.Provider>
  )
}
