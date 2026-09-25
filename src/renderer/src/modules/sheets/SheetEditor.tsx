import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore, type KeyboardEvent } from 'react'
import { closestCenter, DndContext, PointerSensor, useSensor, useSensors, type DragEndEvent } from '@dnd-kit/core'
import { horizontalListSortingStrategy, SortableContext } from '@dnd-kit/sortable'
import DataEditor, {
  CompactSelection,
  GridCellKind,
  type DrawCellCallback,
  type EditableGridCell,
  type GridCell,
  type GridColumn,
  type GridSelection,
  type DataEditorRef,
  type Item,
  type Rectangle,
} from '@glideapps/glide-data-grid'
import '@glideapps/glide-data-grid/dist/index.css'
import { Plus } from 'lucide-react'
import { toast } from 'sonner'
import type { OpenDocument } from '../../app/documents-store'
import { useViewStore } from '../../app/view-store'
import { FidelitySurface } from '../../ui/FidelitySurface'
import { Tip } from '../../ui/Tip'
import { FindBar } from './FindBar'
import { FormulaBar } from './FormulaBar'
import { cellFontSize, gridTheme, headerHeight, isLightColor, rowHeight } from './grid-theme'
import { columnName } from './model/address'
import { formatCellValue, formatStatistic } from './model/format'
import type { CellRange, FindMatch, Workbook, WorkbookSheet } from './model/Workbook'
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

function mapCell([col, row]: Item, columns: readonly number[], rows: readonly number[]): Item {
  return [columns[col] ?? 0, rows[row] ?? 0]
}

function mapRange(range: Rectangle, columns: readonly number[], rows: readonly number[]): CellRange {
  const first = mapCell([range.x, range.y], columns, rows)
  const last = mapCell([range.x + Math.max(1, range.width) - 1, range.y + Math.max(1, range.height) - 1], columns, rows)
  return {
    x: Math.min(first[0], last[0]),
    y: Math.min(first[1], last[1]),
    width: Math.abs(last[0] - first[0]) + 1,
    height: Math.abs(last[1] - first[1]) + 1,
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

/** Glide has no underline; draw it under the rendered text using the cell's own font. */
const drawUnderline: DrawCellCallback = (args, drawContent) => {
  drawContent()
  const { cell, ctx, rect, theme } = args
  if (cell.kind !== GridCellKind.Text || !cell.themeOverride?.baseFontStyle?.includes('underline')) return
  const text = cell.displayData
  if (text === '') return
  ctx.save()
  ctx.font = `${cell.themeOverride.baseFontStyle.replace('underline ', '')} ${theme.fontFamily}`
  const width = Math.min(ctx.measureText(text).width, rect.width - theme.cellHorizontalPadding * 2)
  const align = cell.contentAlign ?? 'left'
  const x = align === 'right'
    ? rect.x + rect.width - theme.cellHorizontalPadding - width
    : align === 'center' ? rect.x + (rect.width - width) / 2 : rect.x + theme.cellHorizontalPadding
  const y = Math.round(rect.y + rect.height / 2 + cellFontSize * 0.5) + 0.5
  ctx.strokeStyle = cell.themeOverride.textDark ?? theme.textDark
  ctx.lineWidth = 1
  ctx.beginPath()
  ctx.moveTo(x, y)
  ctx.lineTo(x + width, y)
  ctx.stroke()
  ctx.restore()
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

  const getCell = useCallback((viewCell: Item, rows: readonly number[]): GridCell => {
    const cell = mapCell(viewCell, visibleColumns, rows)
    const value = workbook.getCellValue(sheetId, cell)
    const style = workbook.getStyle(sheetId, cell)
    const displayed = formatCellValue(value, style)
    const font = `${style.underline === true ? 'underline ' : ''}${style.italic === true ? 'italic ' : ''}${style.bold === true ? '600 ' : ''}${cellFontSize}px`
    const textColor = style.textColor ?? (style.fillColor !== undefined && isLightColor(style.fillColor) ? '#1d1f23' : undefined)
    return {
      kind: GridCellKind.Text,
      allowOverlay: true,
      data: workbook.getCellInput(sheetId, cell),
      displayData: displayed,
      copyData: displayed,
      contentAlign: style.align ?? (typeof value === 'number' ? 'right' : 'left'),
      themeOverride: {
        baseFontStyle: font,
        ...(textColor === undefined ? {} : { textDark: textColor }),
        ...(style.fillColor === undefined ? {} : { bgCell: style.fillColor }),
      },
    }
  }, [sheetId, revision, visibleColumns, workbook])

  const editCell = useCallback((viewCell: Item, value: EditableGridCell, rows: readonly number[]) => {
    const cell = mapCell(viewCell, visibleColumns, rows)
    if (value.kind === GridCellKind.Text) workbook.setCellInput(sheetId, cell, value.data)
    if (value.kind === GridCellKind.Number) workbook.setCellInput(sheetId, cell, value.data ?? null)
  }, [sheetId, visibleColumns, workbook])

  const getScrollingCell = useCallback((cell: Item) => getCell(cell, scrollingRows), [getCell, scrollingRows])
  const getFrozenCell = useCallback((cell: Item) => getCell(cell, frozenRows), [frozenRows, getCell])
  const editScrollingCell = useCallback((cell: Item, value: EditableGridCell) => editCell(cell, value, scrollingRows), [editCell, scrollingRows])
  const editFrozenCell = useCallback((cell: Item, value: EditableGridCell) => editCell(cell, value, frozenRows), [editCell, frozenRows])
  // Row markers show workbook row numbers, so a hidden row leaves a gap (1, 2, 4) as in Excel and Numbers.
  const scrollingRowNumber = useCallback((row: number) => (scrollingRows[row] ?? 0) + 1, [scrollingRows])
  const frozenRowNumber = useCallback((row: number) => (frozenRows[row] ?? 0) + 1, [frozenRows])
  const markerWidth = rowMarkerWidth((visibleRows.at(-1) ?? 0) + 1)

  const selectedRows = selection.rows.toArray()
  const selectedColumns = selection.columns.toArray()
  const viewRange = selectedRows.length > 0
    ? { x: 0, y: Math.min(...selectedRows), width: Math.max(1, visibleColumns.length), height: Math.max(...selectedRows) - Math.min(...selectedRows) + 1 }
    : selectedColumns.length > 0
      ? { x: Math.min(...selectedColumns), y: 0, width: Math.max(...selectedColumns) - Math.min(...selectedColumns) + 1, height: Math.max(1, visibleRows.length) }
      : selection.current?.range ?? { x: 0, y: 0, width: 1, height: 1 }
  const range = mapRange(viewRange, visibleColumns, visibleRows)
  const selectedCell = mapCell(selection.current?.cell ?? [0, 0], visibleColumns, visibleRows)

  const controller = useMemo<SheetController>(() => ({
    documentId: document.id,
    workbook,
    sheetId,
    sheetName: activeSheet.name,
    range,
    cell: selectedCell,
    openFind: () => setFindOpen(true),
  // Range and cell are fresh objects each render, so the memo keys on their contents.
  }), [document.id, workbook, sheetId, activeSheet.name, range.x, range.y, range.width, range.height, selectedCell[0], selectedCell[1]])

  useEffect(() => setActiveSheet(controller), [controller])

  const stats = workbook.selectionStats(sheetId, range)
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
    if (event.key === 'Escape' && findOpen) setFindOpen(false)
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
    rowHeight,
    drawCell: drawUnderline,
    freezeColumns: workbook.getFrozenColumnCount(sheetId),
    // Glide would clear cell by cell (one undo step each); clear the whole selection as one step instead.
    onDelete: () => {
      workbook.clearContents(sheetId, range)
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
            <div className="sheet-frozen-rows" style={{ height: headerHeight + frozenRowCount * rowHeight + 1 }}>
              <DataEditor
                {...sharedGridProps}
                ref={frozenGridRef}
                headerHeight={headerHeight}
                rows={frozenRows.length}
                getCellContent={getFrozenCell}
                onCellEdited={editFrozenCell}
                gridSelection={selectionForRows(selection, 0, frozenRowCount)}
                onGridSelectionChange={next => setSelection(selectionFromRows(next, 0))}
                onPaste={(target, values) => {
                  workbook.setBlock(sheetId, mapCell(target, visibleColumns, frozenRows), values)
                  return false
                }}
                onFillPattern={event => {
                  event.preventDefault()
                  workbook.fillRange(sheetId, mapRange(event.patternSource, visibleColumns, frozenRows), mapRange(event.fillDestination, visibleColumns, frozenRows))
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
              onCellEdited={editScrollingCell}
              gridSelection={selectionForRows(selection, frozenRowCount, scrollingRows.length)}
              onGridSelectionChange={next => setSelection(selectionFromRows(next, frozenRowCount))}
              onPaste={(target, values) => {
                workbook.setBlock(sheetId, mapCell(target, visibleColumns, scrollingRows), values)
                return false
              }}
              onFillPattern={event => {
                event.preventDefault()
                workbook.fillRange(sheetId, mapRange(event.patternSource, visibleColumns, scrollingRows), mapRange(event.fillDestination, visibleColumns, scrollingRows))
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
                      const renamed = workbook.renameSheet(sheet.id, name)
                      if (!renamed) toast.error('That sheet name is already used or not allowed.')
                      else if (sheet.name === activeSheet.name) setActiveSheetName(name.trim())
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
                  <Plus aria-hidden="true" size={14} strokeWidth={1.8} />
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
