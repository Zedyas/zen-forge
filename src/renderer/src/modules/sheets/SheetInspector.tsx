import { Trash2 } from 'lucide-react'
import { useState } from 'react'
import { toast } from 'sonner'
import { cellName, columnName } from './model/address'
import { formatStatistic } from './model/format'
import type { CellRange } from './model/Workbook'
import type { SheetController } from './sheet-commands'
import { Inspector, InspectorRow, InspectorSection } from '../../ui/Inspector'
import { Tip } from '../../ui/Tip'
import { ToolButton } from '../../ui/Toolbar'

function rangeLabel(range: CellRange): string {
  const start = cellName([range.x, range.y])
  const end = cellName([range.x + range.width - 1, range.y + range.height - 1])
  return start === end ? start : `${start}:${end}`
}

/** Detail that does not need to be one click away: statistics, layout, and named ranges. */
export function SheetInspector({ controller }: { readonly controller: SheetController }) {
  const { workbook, sheetId, range, cell } = controller
  const [rangeName, setRangeName] = useState('')
  const stats = workbook.selectionStats(sheetId, range)
  const style = workbook.getStyle(sheetId, cell)
  const frozenRows = workbook.getFrozenRowCount(sheetId)
  const frozenColumns = workbook.getFrozenColumnCount(sheetId)
  const hiddenRows = workbook.getRowCount(sheetId) - workbook.visibleRows(sheetId).length
  const hiddenColumns = workbook.getColumnCount(sheetId) - workbook.visibleColumns(sheetId).length

  const addName = (): void => {
    if (!workbook.addNamedRange(rangeName, sheetId, range)) {
      toast.error('Use a unique name that starts with a letter or underscore, without spaces.')
      return
    }
    setRangeName('')
  }

  return (
    <Inspector label="Spreadsheet inspector">
      <InspectorSection title={`Selection  ${rangeLabel(range)}`}>
        <InspectorRow label="Cells with values"><span>{stats.count.toLocaleString()}</span></InspectorRow>
        {stats.numericCount > 0 && (
          <>
            <InspectorRow label="Sum"><span>{formatStatistic(stats.sum, style)}</span></InspectorRow>
            <InspectorRow label="Average"><span>{formatStatistic(stats.sum / stats.numericCount, style)}</span></InspectorRow>
            <InspectorRow label="Minimum"><span>{formatStatistic(stats.min, style)}</span></InspectorRow>
            <InspectorRow label="Maximum"><span>{formatStatistic(stats.max, style)}</span></InspectorRow>
          </>
        )}
      </InspectorSection>

      <InspectorSection title="Number">
        <InspectorRow label="Format"><span>{style.numberFormat === 'general' ? 'General' : `${style.numberFormat.charAt(0).toUpperCase()}${style.numberFormat.slice(1)}`}</span></InspectorRow>
        {style.numberFormat !== 'general' && style.numberFormat !== 'date' && (
          <InspectorRow label="Decimal places"><span>{style.decimalPlaces}</span></InspectorRow>
        )}
      </InspectorSection>

      <InspectorSection title="Freeze">
        <InspectorRow label="Rows">
          <span className="segmented">
            <button type="button" aria-pressed={frozenRows === 0} onClick={() => workbook.freezeRows(sheetId, 0)}>None</button>
            <button type="button" aria-pressed={frozenRows === 1} onClick={() => workbook.freezeRows(sheetId, 1)}>1</button>
            <Tip label={`Freeze rows 1 to ${range.y + range.height}`}>
              <button type="button" aria-pressed={frozenRows > 1} onClick={() => workbook.freezeRows(sheetId, range.y + range.height)}>Selection</button>
            </Tip>
          </span>
        </InspectorRow>
        <InspectorRow label="Columns">
          <span className="segmented">
            <button type="button" aria-pressed={frozenColumns === 0} onClick={() => workbook.freezeColumns(sheetId, 0)}>None</button>
            <button type="button" aria-pressed={frozenColumns === 1} onClick={() => workbook.freezeColumns(sheetId, 1)}>1</button>
            <Tip label={`Freeze columns A to ${columnName(range.x + range.width - 1)}`}>
              <button type="button" aria-pressed={frozenColumns > 1} onClick={() => workbook.freezeColumns(sheetId, range.x + range.width)}>Selection</button>
            </Tip>
          </span>
        </InspectorRow>
      </InspectorSection>

      <InspectorSection title="Sheet">
        <InspectorRow label="Used area"><span>{workbook.usedSize(sheetId).height} rows × {workbook.usedSize(sheetId).width} columns</span></InspectorRow>
        {hiddenRows > 0 && (
          <InspectorRow label={`${hiddenRows} hidden ${hiddenRows === 1 ? 'row' : 'rows'}`}>
            <button type="button" className="button is-quiet" onClick={() => workbook.showAllRows(sheetId)}>Show</button>
          </InspectorRow>
        )}
        {hiddenColumns > 0 && (
          <InspectorRow label={`${hiddenColumns} hidden ${hiddenColumns === 1 ? 'column' : 'columns'}`}>
            <button type="button" className="button is-quiet" onClick={() => workbook.showAllColumns(sheetId)}>Show</button>
          </InspectorRow>
        )}
      </InspectorSection>

      <InspectorSection title="Named ranges">
        <div className="inspector-inline-form">
          <input
            className="text-field"
            value={rangeName}
            placeholder={`Name ${rangeLabel(range)}`}
            aria-label="Name for the selected range"
            onChange={event => setRangeName(event.currentTarget.value)}
            onKeyDown={event => {
              if (event.key === 'Enter') addName()
            }}
          />
          <button type="button" className="button" disabled={rangeName.trim() === ''} onClick={addName}>Add</button>
        </div>
        {workbook.namedRanges().length === 0
          ? <p className="inspector-note">Name a range to use it in formulas, like =SUM(Expenses).</p>
          : (
            <ul className="inspector-list">
              {workbook.namedRanges().map(namedRange => (
                <li key={namedRange.name}>
                  <span><strong>{namedRange.name}</strong> <code>{namedRange.sheetName}!{rangeLabel(namedRange.range)}</code></span>
                  <ToolButton icon={Trash2} label={`Delete ${namedRange.name}`} onClick={() => workbook.removeNamedRange(namedRange.name)} />
                </li>
              ))}
            </ul>
          )}
      </InspectorSection>
    </Inspector>
  )
}
