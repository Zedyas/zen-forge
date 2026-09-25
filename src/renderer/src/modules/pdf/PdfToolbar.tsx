import {
  Eraser,
  FileOutput,
  FilePlus2,
  Highlighter,
  MousePointer2,
  Palette,
  PenLine,
  Redo2,
  RotateCcw,
  RotateCw,
  Search,
  Square,
  SquareDashed,
  Trash2,
  Type,
  Undo2,
  ZoomIn,
  ZoomOut,
  type LucideIcon,
} from 'lucide-react'
import type { CommandId } from '@shared/commands'
import { ColorTool, ToolButton, Toolbar, ToolSeparator } from '../../ui/Toolbar'
import type { PdfColor } from './engine'
import { useFindStore } from './find-store'
import { findMarkup, updateMarkup } from './markup-actions'
import type { Markup } from './model'
import { runPdfCommand, setZoom } from './pdf-actions'
import { fitWidthZoom } from './PdfDesk'
import { readyPdf, type ReadyPdf } from './pdf-store'
import { SignatureTool } from './SignatureTool'
import { highlightColors, inkColors, setTool, toolKeys, useToolStore, type Tool } from './tool-store'

interface ToolDefinition {
  readonly tool: Tool
  readonly icon: LucideIcon
  readonly label: string
}

const tools: readonly ToolDefinition[] = [
  { tool: 'select', icon: MousePointer2, label: 'Select and move' },
  { tool: 'text', icon: Type, label: 'Text' },
  { tool: 'highlight', icon: Highlighter, label: 'Highlight' },
  { tool: 'ink', icon: PenLine, label: 'Draw' },
  { tool: 'rect', icon: Square, label: 'Rectangle' },
  { tool: 'whiteout', icon: Eraser, label: 'White-out' },
]

const redactTool: ToolDefinition = { tool: 'redact', icon: SquareDashed, label: 'Redact' }

function run(command: CommandId): void {
  void runPdfCommand(command)
}

function ToolChoice({ definition }: { readonly definition: ToolDefinition }) {
  const active = useToolStore(state => state.tool === definition.tool)
  return (
    <ToolButton
      icon={definition.icon}
      label={definition.label}
      shortcut={toolKeys[definition.tool]}
      pressed={active}
      onClick={() => setTool(definition.tool)}
    />
  )
}

/** Colour applies to the selected markup when there is one, and to the next one drawn. */
function recolor(documentId: string, color: PdfColor, highlight: boolean): void {
  useToolStore.setState(highlight ? { highlightColor: color } : { inkColor: color })
  const { selection } = useToolStore.getState()
  const document = readyPdf(documentId)
  if (selection === undefined || document === undefined) return
  const markup = findMarkup(document.present, selection.pageKey, selection.markupId)
  if (markup === undefined) return
  const next: Markup | undefined = markup.kind === 'text' || markup.kind === 'ink'
    ? { ...markup, color }
    : markup.kind === 'rect' && markup.stroke !== undefined
      ? { ...markup, stroke: color }
      : markup.kind === 'rect' && markup.opacity < 1
        ? { ...markup, fill: color }
        : undefined
  if (next !== undefined) updateMarkup(documentId, selection.pageKey, selection.markupId, next)
}

function isColor(value: string | undefined): value is PdfColor {
  return value !== undefined && value.startsWith('#')
}

function ColorChoice({ documentId }: { readonly documentId: string }) {
  const tool = useToolStore(state => state.tool)
  const inkColor = useToolStore(state => state.inkColor)
  const highlightColor = useToolStore(state => state.highlightColor)
  const selectedHighlight = useToolStore(state => {
    const document = readyPdf(documentId)
    if (state.selection === undefined || document === undefined) return false
    const markup = findMarkup(document.present, state.selection.pageKey, state.selection.markupId)
    return markup?.kind === 'rect' && markup.fill !== undefined && markup.opacity < 1
  })
  const highlight = tool === 'highlight' || selectedHighlight
  return (
    <ColorTool
      icon={Palette}
      label={highlight ? 'Highlight colour' : 'Ink colour'}
      options={highlight ? highlightColors : inkColors}
      current={highlight ? highlightColor : inkColor}
      onApply={value => { if (isColor(value)) recolor(documentId, value, highlight) }}
      onChoose={value => { if (isColor(value)) recolor(documentId, value, highlight) }}
    />
  )
}

function ZoomLabel({ documentId, document }: { readonly documentId: string; readonly document: ReadyPdf }) {
  return (
    <ToolButton
      label="Zoom: click to switch between fit width and actual size"
      text={`${Math.round(document.zoom * 100)}%`}
      onClick={() => {
        const fit = fitWidthZoom(document)
        setZoom(documentId, fit !== undefined && Math.abs(document.zoom - 1) < 0.01 ? fit : 1)
      }}
    />
  )
}

export function PdfToolbar({ documentId, document }: { readonly documentId: string; readonly document: ReadyPdf }) {
  const canUndo = document.past.length > 0
  const canRedo = document.future.length > 0
  const findOpen = useFindStore(find => find.open)
  return (
    <Toolbar label="PDF tools">
      <ToolButton icon={Undo2} label="Undo" command="edit.undo" disabled={!canUndo} onClick={() => run('edit.undo')} />
      <ToolButton icon={Redo2} label="Redo" command="edit.redo" disabled={!canRedo} onClick={() => run('edit.redo')} />
      <ToolSeparator />
      {tools.map(definition => <ToolChoice key={definition.tool} definition={definition} />)}
      <SignatureTool />
      <ToolChoice definition={redactTool} />
      <ToolSeparator />
      <ColorChoice documentId={documentId} />
      <ToolSeparator />
      <ToolButton icon={RotateCcw} label="Rotate page left" command="page.rotateLeft" onClick={() => run('page.rotateLeft')} />
      <ToolButton icon={RotateCw} label="Rotate page right" command="page.rotateRight" onClick={() => run('page.rotateRight')} />
      <ToolButton icon={FilePlus2} label="Insert pages from file…" onClick={() => run('page.insert')} />
      <ToolButton icon={FileOutput} label="Extract page…" onClick={() => run('page.extract')} />
      <ToolButton icon={Trash2} label="Delete page" command="page.delete" onClick={() => run('page.delete')} />
      <span className="toolbar-grow" />
      <ToolButton icon={Search} label="Find" command="edit.find" pressed={findOpen} onClick={() => run('edit.find')} />
      <ToolSeparator />
      <ToolButton icon={ZoomOut} label="Zoom out" command="view.zoomOut" onClick={() => run('view.zoomOut')} />
      <ZoomLabel documentId={documentId} document={document} />
      <ToolButton icon={ZoomIn} label="Zoom in" command="view.zoomIn" onClick={() => run('view.zoomIn')} />
    </Toolbar>
  )
}

/** The tools that stay in the title bar when the toolbar is hidden. */
export function PdfEssentials() {
  const essentials = tools.filter(definition => definition.tool === 'select' || definition.tool === 'text' || definition.tool === 'highlight')
  return (
    <>
      {essentials.map(definition => <ToolChoice key={definition.tool} definition={definition} />)}
      <SignatureTool />
      <ToolButton icon={RotateCw} label="Rotate page right" command="page.rotateRight" onClick={() => run('page.rotateRight')} />
    </>
  )
}
