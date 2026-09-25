import { ContextMenu } from '@base-ui/react/context-menu'
import { useCallback, useEffect, useState } from 'react'
import { FilePlus2, FolderOpen, Files, Images, Search } from 'lucide-react'
import { toast } from 'sonner'
import { applicationForExtension, applications, findApplication, suiteName, type EditorApplicationId } from '@shared/applications'
import type { Appearance } from '@shared/shell'
import { forgetRecent, listRecents, watchRecents, type RecentDocument } from '../services/index/document-index'
import { platformClient } from '../services/platform/client'
import { combinePdfs, createPdfFromImages } from '../modules/pdf/workflows'
import { AppIcon } from '../ui/AppIcon'
import { commandShortcut, Tip } from '../ui/Tip'
import { chooseAndOpenFiles, openFiles } from './document-actions'
import { useDocumentsStore } from './documents-store'
import { UpdateOffer } from './updates'


function reportFailure(title: string) {
  return (error: unknown): void => {
    toast.error(title, { description: error instanceof Error ? error.message : undefined })
  }
}

function openedLabel(openedAt: number): string {
  const opened = new Date(openedAt)
  const today = new Date()
  const days = Math.floor((today.setHours(0, 0, 0, 0) - new Date(openedAt).setHours(0, 0, 0, 0)) / 86_400_000)
  if (days === 0) return `Today, ${opened.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })}`
  if (days === 1) return 'Yesterday'
  if (days < 7) return opened.toLocaleDateString(undefined, { weekday: 'long' })
  return opened.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
}

function AppearanceControl() {
  const [appearance, setAppearance] = useState<Appearance>('system')
  // The View menu can change appearance from any window, so re-read when Home regains focus or the
  // rendered scheme flips.
  useEffect(() => {
    const refresh = (): void => void platformClient.getAppearance().then(setAppearance)
    const scheme = window.matchMedia('(prefers-color-scheme: dark)')
    refresh()
    window.addEventListener('focus', refresh)
    scheme.addEventListener('change', refresh)
    return () => {
      window.removeEventListener('focus', refresh)
      scheme.removeEventListener('change', refresh)
    }
  }, [])
  const options: ReadonlyArray<readonly [Appearance, string]> = [['system', 'Auto'], ['light', 'Light'], ['dark', 'Dark']]
  return (
    <span className="segmented" role="group" aria-label="Appearance">
      {options.map(([value, label]) => (
        <button
          key={value}
          type="button"
          aria-pressed={appearance === value}
          onClick={() => {
            setAppearance(value)
            void platformClient.setAppearance(value)
          }}
        >
          {label}
        </button>
      ))}
    </span>
  )
}

function RecentCard({ file }: { readonly file: RecentDocument }) {
  const application = applicationForExtension(file.extension)
  const kind = application === undefined ? file.extension.toUpperCase() : findApplication(application).name
  return (
    <ContextMenu.Root>
      <ContextMenu.Trigger
        render={<button type="button" />}
        className="recent-card"
        title={file.path}
        onClick={() => void openFiles([file]).catch(reportFailure('Could not open the file'))}
      >
        <span className={`recent-preview${file.preview === undefined ? ' is-placeholder' : ''}`}>
          {file.preview === undefined
            ? application !== undefined && <AppIcon application={application} size={44} />
            : <img src={file.preview} alt="" />}
        </span>
        <span>
          <strong>{file.name.replace(/\.[^.]+$/, '')}</strong>
          <small>{kind}, {openedLabel(file.openedAt)}</small>
        </span>
      </ContextMenu.Trigger>
      <ContextMenu.Portal>
        <ContextMenu.Positioner>
          <ContextMenu.Popup className="menu-popup">
            <ContextMenu.Item className="menu-item" onClick={() => void openFiles([file]).catch(reportFailure('Could not open the file'))}><span>Open</span></ContextMenu.Item>
            <ContextMenu.Item className="menu-item" onClick={() => void platformClient.revealInFinder(file.path)}><span>Show in Finder</span></ContextMenu.Item>
            <ContextMenu.Separator className="menu-separator" />
            <ContextMenu.Item className="menu-item" onClick={() => void forgetRecent(file.path)}><span>Remove from Recents</span></ContextMenu.Item>
          </ContextMenu.Popup>
        </ContextMenu.Positioner>
      </ContextMenu.Portal>
    </ContextMenu.Root>
  )
}

/** The Home tab: start something, reopen something. Files can be dropped anywhere in the window. */
export function HomeScreen() {
  const [query, setQuery] = useState('')
  const [recents, setRecents] = useState<readonly RecentDocument[]>([])

  const refresh = useCallback(() => {
    void listRecents(query).then(setRecents)
  }, [query])

  useEffect(() => {
    refresh()
    return watchRecents(refresh)
  }, [refresh])

  const start = (application: EditorApplicationId): void => {
    if (application === 'sheets') useDocumentsStore.getState().openUntitled('sheets')
    else void chooseAndOpenFiles(application).catch(reportFailure('Could not open the file'))
  }

  return (
    <main className="home">
      <div className="home-body">
        <UpdateOffer />
        <h1>{suiteName}</h1>
        <p className="home-lede">Spreadsheets and PDFs, kept on this Mac.</p>

        <div className="home-apps">
          {applications.map(application => {
            const card = (
              <button
                key={application.id}
                type="button"
                className="home-app"
                disabled={!application.available}
                onClick={() => start(application.id)}
              >
                <AppIcon application={application.id} size={40} unavailable={!application.available} />
                <span>
                  <strong>{application.name}</strong>
                  <small>{application.available ? application.kind : `${application.kind}, planned`}</small>
                </span>
              </button>
            )
            if (application.id === 'sheets') return <Tip key={application.id} label="New spreadsheet" shortcut={commandShortcut('file.new')}>{card}</Tip>
            if (application.id === 'pdf') return <Tip key={application.id} label="Open a PDF">{card}</Tip>
            return card
          })}
        </div>

        <div className="home-actions">
          <button type="button" className="button" onClick={() => start('sheets')}>
            <FilePlus2 aria-hidden="true" size={16} />New spreadsheet
          </button>
          <button type="button" className="button" onClick={() => void chooseAndOpenFiles().catch(reportFailure('Could not open the file'))}>
            <FolderOpen aria-hidden="true" size={16} />Open…
          </button>
          <button type="button" className="button" onClick={() => void createPdfFromImages().catch(reportFailure('Could not create the PDF'))}>
            <Images aria-hidden="true" size={16} />PDF from images…
          </button>
          <button type="button" className="button" onClick={() => void combinePdfs().catch(reportFailure('Could not combine the PDFs'))}>
            <Files aria-hidden="true" size={16} />Combine PDFs…
          </button>
        </div>

        <div className="home-section-head">
          <h2>Recent</h2>
          <label className="home-search">
            <Search aria-hidden="true" size={15} />
            <input type="search" value={query} placeholder="Search recents" aria-label="Search recent files" onChange={event => setQuery(event.currentTarget.value)} />
          </label>
        </div>
        {recents.length === 0 ? (
          <div className="home-empty">
            <strong>{query === '' ? 'Nothing opened yet' : 'No recent file matches'}</strong>
            {query === '' ? 'Files you open or save appear here. Drop a spreadsheet or PDF anywhere in this window to open it.' : 'Try another part of the file name.'}
          </div>
        ) : (
          <div className="recent-grid">
            {recents.map(file => <RecentCard key={file.path} file={file} />)}
          </div>
        )}
      </div>

      <footer className="home-footer">
        <span>Files stay on this Mac. No account needed.</span>
        <AppearanceControl />
      </footer>
    </main>
  )
}
