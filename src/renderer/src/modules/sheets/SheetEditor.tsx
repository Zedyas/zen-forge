import { useCallback, useEffect, useEffectEvent, useMemo, useRef, useState, useSyncExternalStore, type KeyboardEvent } from 'react'
import { closestCenter, DndContext, PointerSensor, useSensor, useSensors, type DragEndEvent } from '@dnd-kit/core'
import { horizontalListSortingStrategy, SortableContext } from '@dnd-kit/sortable'
import DataEditor, {
  CompactSelection,
  GridCellKind,
  type DrawCellCallback,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type GridKeyEventArgs,
  type GridSelection,
  type DataEditorRef,
  type Highlight,
  type Item,
  type Rectangle,
} from '@glideapps/glide-data-grid'
import '@glideapps/glide-data-grid/dist/index.css'
import { Plus } from 'lucide-react'
import { toast } from 'sonner'
import type { OpenDocument } from '../../app/documents-store'
import { useViewStore } from '../../app/view-store'
import { FidelitySurface } from '../../ui/FidelitySurface'
import { smallIcon } from '../../ui/icons'
import { Tip } from '../../ui/Tip'
import { cancelCut, copyCells, pasteCells, usePendingCut } from './clipboard'
import { FindBar } from './FindBar'
import { FormulaBar } from './FormulaBar'
import { drawCellText, drawsOwnText, fontStack, fontStyle, measureCellHeight } from './cell-text'
import { gridTheme, headerHeight, isLightColor } from './grid-theme'
import { coverBlocks, fillLines, mapCell, mapRange, snapCell, viewRange as viewRangeOf, visibleAreas } from './grid-mapping'
import { columnName, sheetNameProblem } from './model/address'
import { formatCellValue, formatColor, formatStatistic } from './model/format'
import type { FindMatch, Workbook, WorkbookSheet } from './model/Workbook'
import { setActiveSheet, type SheetController } from './sheet-commands'
import { SheetContextMenu } from './SheetContextMenu'
import { SheetInspector } from './SheetInspector'
import { SheetTab } from './SheetTab'
import { SheetToolbar } from './SheetToolbar'
import './sheets.css'

function selectionForCell(cell: Item): GridSelection {
  return {
    current: { cell, range: { x: cell[0], y: cell[1], width: 1, height: 1 }, rangeStack: [] },
    columns: CompactSelection.empty(),
    rows: CompactSelection.empty(),
  }
}

function requiredSheet(sheets: readonly WorkbookSheet[], name: string): WorkbookSheet {
  const sheet = sheets.find(candidate => candidate.name === name) ?? sheets[0]
  if (sheet === undefined) throw new Error('A workbook must contain a sheet.')
  return sheet
}

/* The frozen-row area is a second grid stacked above the scrolling one; these translate one
   shared selection into each grid's own row coordinates and back. */

function rowsInSegment(selection: CompactSelection, rowOffset: number, rowCount: number): CompactSelection {
  let result = CompactSelection.empty()
  selection.toArray().forEach(row => {
    if (row >= rowOffset && row < rowOffset + rowCount) result = result.add(row - rowOffset)
  })
  return result
}

function selectionForRows(selection: GridSelection, rowOffset: number, rowCount: number): GridSelection {
  const current = selection.current
  const rows = rowsInSegment(selection.rows, rowOffset, rowCount)
  if (current === undefined || current.cell[1] < rowOffset || current.cell[1] >= rowOffset + rowCount) {
    return { columns: selection.columns, rows }
  }
  const rangeStart = Math.max(current.range.y, rowOffset)
  const rangeEnd = Math.min(current.range.y + current.range.height, rowOffset + rowCount)
  return {
    current: {
      cell: [current.cell[0], current.cell[1] - rowOffset],
      range: { ...current.range, y: rangeStart - rowOffset, height: Math.max(1, rangeEnd - rangeStart) },
      rangeStack: current.rangeStack.flatMap(range => {
        const start = Math.max(range.y, rowOffset)
        const end = Math.min(range.y + range.height, rowOffset + rowCount)
        return end > start ? [{ ...range, y: start - rowOffset, height: end - start }] : []
      }),
    },
    columns: selection.columns,
    rows,
  }
}

/** A view rectangle in one grid's own rows, or undefined when none of it is in that grid. */
function rectangleInRows(rectangle: Rectangle, rowOffset: number, rowCount: number): Rectangle | undefined {
  const start = Math.max(rectangle.y, rowOffset)
  const end = Math.min(rectangle.y + rectangle.height, rowOffset + rowCount)
  return end > start ? { ...rectangle, y: start - rowOffset, height: end - start } : undefined
}

function selectionFromRows(selection: GridSelection, rowOffset: number): GridSelection {
  const current = selection.current
  return {
    current: current === undefined ? undefined : {
      cell: [current.cell[0], current.cell[1] + rowOffset],
      range: { ...current.range, y: current.range.y + rowOffset },
      rangeStack: current.rangeStack.map(range => ({ ...range, y: range.y + rowOffset })),
    },
    columns: selection.columns,
    rows: selection.rows.offset(rowOffset),
  }
}

/**
 * Glide sizes the row-number column from its own row count, so the frozen grid (a few rows) would get a
 * narrower one than the grid below it and their columns would not line up. Both use the width Glide
 * picks for the largest row number shown.
 */
function rowMarkerWidth(largestRowNumber: number): number {
  return largestRowNumber > 10_000 ? 48 : largestRowNumber > 1_000 ? 44 : largestRowNumber > 100 ? 36 : 32
}

/** Recomputes the grid colours when macOS switches between light and dark. */
function useGridTheme(): ReturnType<typeof gridTheme> {
  const [version, setVersion] = useState(0)
  useEffect(() => {
    const query = window.matchMedia('(prefers-color-scheme: dark)')
    const update = (): void => setVersion(value => value + 1)
    query.addEventListener('change', update)
    return () => query.removeEventListener('change', update)
  }, [])
  // `version` is only the recompute trigger; gridTheme reads the live CSS variables.
  return useMemo(() => gridTheme(), [version])
}

interface SheetEditorProps {
  readonly document: OpenDocument
  readonly workbook: Workbook
}

/** Glide adapter that keeps view coordinates (hidden rows removed) separate from workbook coordinates. */
export function SheetEditor({ document, workbook }: SheetEditorProps) {
  const revision = useSyncExternalStore(workbook.subscribe, workbook.getSnapshot)
  const theme = useGridTheme()
  const inspectorShown = useViewStore(state => state.inspector)
  const sheets = workbook.sheets()
  const [activeSheetName, setActiveSheetName] = useState(sheets[0]?.name ?? 'Sheet1')
  const [selection, setSelection] = useState<GridSelection>(() => selectionForCell([0, 0]))
  const [findOpen, setFindOpen] = useState(false)
  const frozenGridRef = useRef<DataEditorRef>(null)
  const activeSheet = requiredSheet(sheets, activeSheetName)
  const sheetId = activeSheet.id

  // `revision` is the change signal: the workbook mutates in place, so its identity never changes.
  const visibleRows = useMemo(() => workbook.visibleRows(sheetId), [sheetId, revision, workbook])
  const visibleColumns = useMemo(() => workbook.visibleColumns(sheetId), [sheetId, revision, workbook])
  const frozenRowCount = workbook.getFrozenRowCount(sheetId)
  const frozenRows = useMemo(() => visibleRows.slice(0, frozenRowCount), [frozenRowCount, visibleRows])
  const scrollingRows = useMemo(() => frozenRowCount === 0 ? visibleRows : visibleRows.slice(frozenRowCount), [frozenRowCount, visibleRows])
  const sheetSensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 5 } }))

  const columns = useMemo<readonly GridColumn[]>(() => visibleColumns.map(modelColumn => ({
    id: columnName(modelColumn),
    title: columnName(modelColumn),
    width: workbook.getColumnWidth(sheetId, modelColumn),
  })), [sheetId, revision, visibleColumns, workbook])

  // Merged blocks, each with where it shows in the grid (view coordinates across both grids).
  const merges = useMemo(() => workbook.merges(sheetId).flatMap(model => {
    const view = viewRangeOf(model, visibleColumns, visibleRows)
    return view === undefined ? [] : [{ model, view }]
  }), [sheetId, revision, visibleColumns, visibleRows, workbook])
  const mergeOf = useCallback((cell: Item) => {
    const model = workbook.mergeAt(sheetId, cell)
    return model === undefined ? undefined : merges.find(merge => merge.model === model)
  }, [merges, sheetId, workbook])

  /**
   * Every cell of a merged block shows the block's top-left cell. Each row of the block is one span
   * across its columns; `spanBelow` (patched into Glide) leaves out the grid line between its rows.
   */
  const getCell = useCallback((viewCell: Item, rows: readonly number[]): GridCell => {
    const merge = mergeOf(mapCell(viewCell, visibleColumns, rows))
    const cell: Item = merge === undefined ? mapCell(viewCell, visibleColumns, rows) : [merge.model.x, merge.model.y]
    const value = workbook.getCellValue(sheetId, cell)
    const style = workbook.getStyle(sheetId, cell)
    const displayed = formatCellValue(value, style)
    const textColor = formatColor(value, style) ?? style.textColor ?? (style.fillColor !== undefined && isLightColor(style.fillColor) ? '#1d1f23' : undefined)
    return {
      kind: GridCellKind.Text,
      allowOverlay: true,
      data: workbook.getCellInput(sheetId, cell),
      displayData: displayed,
      copyData: displayed,
      contentAlign: style.align ?? (typeof value === 'number' ? 'right' : 'left'),
      // The cell editor then wraps too, and takes Return for a new line with ⌥.
      allowWrapping: style.wrap === true,
      ...(merge === undefined ? {} : {
        span: [merge.view.x, merge.view.x + merge.view.width - 1] as const,
        spanBelow: (rows[viewCell[1] + 1] ?? Number.POSITIVE_INFINITY) < merge.model.y + merge.model.height,
      }),
      themeOverride: {
        baseFontStyle: fontStyle(style),
        ...(style.fontFamily === undefined || theme.fontFamily === undefined ? {} : { fontFamily: fontStack(style.fontFamily, theme.fontFamily) }),
        ...(textColor === undefined ? {} : { textDark: textColor }),
        ...(style.fillColor === undefined ? {} : { bgCell: style.fillColor }),
      },
    }
  }, [mergeOf, sheetId, revision, theme, visibleColumns, workbook])

  const editCell = useCallback((viewCell: Item, value: EditableGridCell, rows: readonly number[]) => {
    const merge = mergeOf(mapCell(viewCell, visibleColumns, rows))
    const cell: Item = merge === undefined ? mapCell(viewCell, visibleColumns, rows) : [merge.model.x, merge.model.y]
    if (value.kind === GridCellKind.Text) workbook.setCellInput(sheetId, cell, value.data, measureCellHeight)
    if (value.kind === GridCellKind.Number) workbook.setCellInput(sheetId, cell, value.data ?? null, measureCellHeight)
  }, [mergeOf, sheetId, visibleColumns, workbook])

  /**
   * Glide draws plain text. Wrapped, underlined and struck-through text Ledger draws itself (cell-text.ts),
   * and so is a merged block's, laid out across the whole block while each row draws its own slice.
   */
  const drawCell = useCallback((args: Parameters<DrawCellCallback>[0], drawContent: () => void, rows: readonly number[]) => {
    const merge = mergeOf(mapCell([args.col, args.row], visibleColumns, rows))
    const style = workbook.getStyle(sheetId, merge === undefined ? mapCell([args.col, args.row], visibleColumns, rows) : [merge.model.x, merge.model.y])
    if (args.cell.kind !== GridCellKind.Text || (merge === undefined && !drawsOwnText(style))) {
      drawContent()
      return
    }
    let { y: top, height } = args.rect
    if (merge !== undefined) {
      const inBlock = (index: number): boolean => {
        const row = rows[index]
        return row !== undefined && row >= merge.model.y && row < merge.model.y + merge.model.height
      }
      for (let index = args.row - 1; inBlock(index); index -= 1) {
        const above = workbook.getRowHeight(sheetId, rows[index] ?? 0)
        top -= above
        height += above
      }
      for (let index = args.row + 1; inBlock(index); index += 1) height += workbook.getRowHeight(sheetId, rows[index] ?? 0)
    }
    drawCellText(args, args.cell.displayData, style, { ...args.rect, y: top, height })
  }, [mergeOf, sheetId, revision, visibleColumns, workbook])
  const drawScrollingCell = useCallback<DrawCellCallback>((args, draw) => drawCell(args, draw, scrollingRows), [drawCell, scrollingRows])
  const drawFrozenCell = useCallback<DrawCellCallback>((args, draw) => drawCell(args, draw, frozenRows), [drawCell, frozenRows])
  const scrollingRowHeight = useCallback((row: number) => workbook.getRowHeight(sheetId, scrollingRows[row] ?? 0), [sheetId, revision, scrollingRows, workbook])
  const frozenRowHeight = useCallback((row: number) => workbook.getRowHeight(sheetId, frozenRows[row] ?? 0), [sheetId, revision, frozenRows, workbook])
  const frozenHeight = frozenRows.reduce((total, row) => total + workbook.getRowHeight(sheetId, row), 0)

  const getScrollingCell = useCallback((cell: Item) => getCell(cell, scrollingRows), [getCell, scrollingRows])
  const getFrozenCell = useCallback((cell: Item) => getCell(cell, frozenRows), [frozenRows, getCell])
  const editScrollingCell = useCallback((cell: Item, value: EditableGridCell) => editCell(cell, value, scrollingRows), [editCell, scrollingRows])
  const editFrozenCell = useCallback((cell: Item, value: EditableGridCell) => editCell(cell, value, frozenRows), [editCell, frozenRows])
  // Row markers show workbook row numbers, so a hidden row leaves a gap (1, 2, 4) as in Excel and Numbers.
  const scrollingRowNumber = useCallback((row: number) => (scrollingRows[row] ?? 0) + 1, [scrollingRows])
  const frozenRowNumber = useCallback((row: number) => (frozenRows[row] ?? 0) + 1, [frozenRows])
  const markerWidth = rowMarkerWidth((visibleRows.at(-1) ?? 0) + 1)

  /** Set by an arrow, Tab or Return key, so the selection change it causes steps out of a merged block. */
  const stepped = useRef(false)
  const snapToMerges = (next: GridSelection): GridSelection => {
    const current = next.current
    const blocks = merges.map(merge => merge.view)
    if (current === undefined || blocks.length === 0) return next
    const cell = snapCell(current.cell, selection.current?.cell, stepped.current, blocks)
    stepped.current = false
    const single = current.range.width === 1 && current.range.height === 1
    const range = coverBlocks(single ? { x: cell[0], y: cell[1], width: 1, height: 1 } : current.range, blocks)
    return { ...next, current: { ...current, cell, range } }
  }

  const selectedRows = selection.rows.toArray()
  const selectedColumns = selection.columns.toArray()
  const viewRange = selectedRows.length > 0
    ? { x: 0, y: Math.min(...selectedRows), width: Math.max(1, visibleColumns.length), height: Math.max(...selectedRows) - Math.min(...selectedRows) + 1 }
    : selectedColumns.length > 0
      ? { x: Math.min(...selectedColumns), y: 0, width: Math.max(...selectedColumns) - Math.min(...selectedColumns) + 1, height: Math.max(1, visibleRows.length) }
      : selection.current?.range ?? { x: 0, y: 0, width: 1, height: 1 }
  const range = mapRange(viewRange, visibleColumns, visibleRows)
  // Cell-by-cell actions (clear, format, statistics) skip hidden rows and columns inside the selection.
  const areas = useMemo(
    () => visibleAreas(viewRange, visibleColumns, visibleRows),
    [viewRange.x, viewRange.y, viewRange.width, viewRange.height, visibleColumns, visibleRows],
  )
  const selectedCell = mapCell(selection.current?.cell ?? [0, 0], visibleColumns, visibleRows)

  const controller = useMemo<SheetController>(() => ({
    documentId: document.id,
    workbook,
    sheetId,
    sheetName: activeSheet.name,
    range,
    areas,
    cell: selectedCell,
    openFind: () => setFindOpen(true),
  // Range and cell are fresh objects each render, so the memo keys on their contents.
  }), [document.id, workbook, sheetId, activeSheet.name, range.x, range.y, range.width, range.height, areas, selectedCell[0], selectedCell[1]])

  useEffect(() => setActiveSheet(controller), [controller])

  const stats = workbook.selectionStats(sheetId, areas)
  const cutRange = usePendingCut(workbook, sheetId)
  const cutView = cutRange === undefined ? undefined : viewRangeOf(cutRange, visibleColumns, visibleRows)
  const cutOutline = (rowOffset: number, rowCount: number): readonly Highlight[] | undefined => {
    const rectangle = cutView === undefined ? undefined : rectangleInRows(cutView, rowOffset, rowCount)
    return rectangle === undefined ? undefined : [{ color: theme.accentLight ?? 'rgba(0, 0, 0, 0.08)', range: rectangle, style: 'dashed' }]
  }

  // Edit > Copy, Cut and Paste arrive as clipboard events; they act on the grid only while it has focus.
  const handleClipboard = useEffectEvent((event: ClipboardEvent) => {
    const focused = window.document.activeElement
    if (!(focused instanceof HTMLElement) || focused.closest('.sheet-canvas') === null || event.clipboardData === null) return
    if (focused.matches('input, textarea, [contenteditable="true"]')) return
    if (event.type === 'paste') {
      if (pasteCells(event.clipboardData, { workbook, sheetId, selection: viewRange, columns: visibleColumns, rows: visibleRows })) event.preventDefault()
      return
    }
    event.preventDefault()
    copyCells(event.clipboardData, {
      workbook,
      sheetId,
      columns: visibleColumns.slice(viewRange.x, viewRange.x + viewRange.width),
      rows: visibleRows.slice(viewRange.y, viewRange.y + viewRange.height),
      range,
    }, event.type === 'cut')
  })
  useEffect(() => {
    const listener = (event: ClipboardEvent): void => handleClipboard(event)
    const types = ['copy', 'cut', 'paste'] as const
    types.forEach(type => window.document.addEventListener(type, listener))
    return () => types.forEach(type => window.document.removeEventListener(type, listener))
  }, [])
  const activeStyle = workbook.getStyle(sheetId, selectedCell)

  const handleSheetDragEnd = (event: DragEndEvent): void => {
    if (event.over === null) return
    const from = sheets.findIndex(sheet => sheet.id === event.active.id)
    const to = sheets.findIndex(sheet => sheet.id === event.over?.id)
    workbook.moveSheet(from, to)
  }

  const selectMatch = (match: FindMatch): void => {
    setActiveSheetName(match.sheetName)
    const col = workbook.visibleColumns(match.sheetId).indexOf(match.cell[0])
    const row = workbook.visibleRows(match.sheetId).indexOf(match.cell[1])
    if (col < 0 || row < 0) {
      toast.info('That match is in a hidden row or column', { description: 'Show hidden rows and columns from the inspector.' })
      return
    }
    setSelection(selectionForCell([col, row]))
  }

  // ⌘F comes from the native Edit menu; Escape closing find is local to the grid.
  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>): void => {
    if (event.key !== 'Escape') return
    if (findOpen) setFindOpen(false)
    cancelCut()
  }

  /** A right-click outside the selection selects that cell first, as Numbers and Excel do. */
  const selectOnContextMenu = (viewCell: Item, rowOffset: number): void => {
    const [col, row] = [viewCell[0], viewCell[1] + rowOffset]
    const current = selection.current?.range
    const inside = current !== undefined && col >= current.x && col < current.x + current.width
      && row >= current.y && row < current.y + current.height
    if (!inside && col >= 0) setSelection(selectionForCell([col, row]))
  }

  const sharedGridProps = {
    columns,
    getCellsForSelection: true,
    rangeSelect: 'multi-rect',
    columnSelect: 'multi',
    rowSelect: 'multi',
    fillHandle: true,
    editOnType: true,
    smoothScrollX: true,
    width: '100%',
    height: '100%',
    theme,
    onKeyDown: (event: GridKeyEventArgs) => {
      stepped.current = !event.shiftKey && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab', 'Enter'].includes(event.key)
    },
    // Copy, cut and paste are Ledger's own (see clipboard.ts), so they keep formulas and styles.
    keybindings: { copy: false, cut: false, paste: false },
    freezeColumns: workbook.getFrozenColumnCount(sheetId),
    // Glide would clear cell by cell (one undo step each); clear the whole selection as one step instead.
    onDelete: () => {
      workbook.clearContents(sheetId, areas)
      return false
    },
    onColumnResizeEnd: (_column: GridColumn, size: number, viewColumn: number) => {
      const modelColumn = visibleColumns[viewColumn]
      if (modelColumn !== undefined) workbook.setColumnWidth(sheetId, modelColumn, size)
    },
  } as const

  return (
    <div className="window-body" onKeyDownCapture={handleKeyDown}>
      <div className="window-main">
        <SheetToolbar controller={controller} findOpen={findOpen} />
        <FormulaBar workbook={workbook} sheetId={sheetId} selectedCell={selectedCell} />
        {findOpen && <FindBar workbook={workbook} revision={revision} onClose={() => setFindOpen(false)} onSelectMatch={selectMatch} />}
        <SheetContextMenu controller={controller}>
          {frozenRowCount > 0 && (
            <div className="sheet-frozen-rows" style={{ height: headerHeight + frozenHeight + 1 }}>
              <DataEditor
                {...sharedGridProps}
                ref={frozenGridRef}
                headerHeight={headerHeight}
                rows={frozenRows.length}
                getCellContent={getFrozenCell}
                drawCell={drawFrozenCell}
                rowHeight={frozenRowHeight}
                onCellEdited={editFrozenCell}
                gridSelection={selectionForRows(selection, 0, frozenRowCount)}
                onGridSelectionChange={next => setSelection(snapToMerges(selectionFromRows(next, 0)))}
                highlightRegions={cutOutline(0, frozenRowCount)}
                onFillPattern={event => {
                  event.preventDefault()
                  const { lines, backward } = fillLines(event.patternSource, event.fillDestination, visibleColumns, frozenRows)
                  workbook.fillCells(sheetId, lines, backward)
                }}
                rowMarkers={{ kind: 'number', rowNumber: frozenRowNumber, width: markerWidth }}
                onCellContextMenu={cell => selectOnContextMenu(cell, 0)}
              />
            </div>
          )}
          <div className="sheet-scrolling-rows">
            <DataEditor
              {...sharedGridProps}
              headerHeight={frozenRowCount > 0 ? 0 : headerHeight}
              rows={scrollingRows.length}
              getCellContent={getScrollingCell}
              drawCell={drawScrollingCell}
              rowHeight={scrollingRowHeight}
              onCellEdited={editScrollingCell}
              gridSelection={selectionForRows(selection, frozenRowCount, scrollingRows.length)}
              onGridSelectionChange={next => setSelection(snapToMerges(selectionFromRows(next, frozenRowCount)))}
              highlightRegions={cutOutline(frozenRowCount, scrollingRows.length)}
              onFillPattern={event => {
                event.preventDefault()
                const { lines, backward } = fillLines(event.patternSource, event.fillDestination, visibleColumns, scrollingRows)
                workbook.fillCells(sheetId, lines, backward)
              }}
              rowMarkers={{ kind: 'number', rowNumber: scrollingRowNumber, width: markerWidth }}
              onCellContextMenu={cell => selectOnContextMenu(cell, frozenRowCount)}
              smoothScrollY
              onVisibleRegionChanged={region => {
                if (frozenRowCount > 0) frozenGridRef.current?.scrollTo(region.x, 0, 'horizontal')
              }}
            />
          </div>
        </SheetContextMenu>
        <footer className="statusbar">
          <DndContext sensors={sheetSensors} collisionDetection={closestCenter} onDragEnd={handleSheetDragEnd}>
            <div className="sheet-tabs" role="tablist" aria-label="Sheets">
              <SortableContext items={sheets.map(sheet => sheet.id)} strategy={horizontalListSortingStrategy}>
                {sheets.map(sheet => (
                  <SheetTab
                    key={sheet.id}
                    sheet={sheet}
                    active={sheet.name === activeSheet.name}
                    canDelete={sheets.length > 1}
                    onSelect={() => {
                      setActiveSheetName(sheet.name)
                      setSelection(selectionForCell([0, 0]))
                    }}
                    onRename={name => {
                      const problem = sheetNameProblem(name.trim(), sheets.filter(other => other.id !== sheet.id).map(other => other.name))
                      if (problem !== undefined) {
                        toast.error(problem)
                        return false
                      }
                      const renamed = workbook.renameSheet(sheet.id, name)
                      if (renamed && sheet.name === activeSheet.name) setActiveSheetName(name.trim())
                      return renamed
                    }}
                    onDelete={() => {
                      const next = sheets.find(candidate => candidate.id !== sheet.id)
                      workbook.removeSheet(sheet.id)
                      if (next !== undefined && sheet.name === activeSheet.name) setActiveSheetName(next.name)
                    }}
                  />
                ))}
              </SortableContext>
              <Tip label="Add sheet">
                <button className="tool" type="button" aria-label="Add sheet" onClick={() => setActiveSheetName(workbook.addSheet().name)}>
                  <Plus {...smallIcon} />
                </button>
              </Tip>
            </div>
          </DndContext>
          <span className="statusbar-grow" />
          {stats.numericCount > 0 && (
            <>
              <span className="statusbar-item">Sum <b>{formatStatistic(stats.sum, activeStyle)}</b></span>
              <span className="statusbar-item">Average <b>{formatStatistic(stats.sum / stats.numericCount, activeStyle)}</b></span>
            </>
          )}
          {stats.count > 1 && <span className="statusbar-item">Count <b>{stats.count.toLocaleString()}</b></span>}
          <FidelitySurface documentId={document.id} />
        </footer>
      </div>
      {inspectorShown && <SheetInspector controller={controller} />}
    </div>
  )
}
