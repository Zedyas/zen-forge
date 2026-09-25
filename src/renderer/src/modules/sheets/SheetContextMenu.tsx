import { ContextMenu } from '@base-ui/react/context-menu'
import type { ReactNode } from 'react'
import { sortByActiveColumn, type SheetController } from './sheet-commands'

interface SheetContextMenuProps {
  readonly controller: SheetController
  readonly children: ReactNode
}

/** Right-click menu over the grid; acts on the selection the right-click left in place. */
export function SheetContextMenu({ controller, children }: SheetContextMenuProps) {
  const { workbook, sheetId, range } = controller
  const rows = `${range.height} ${range.height === 1 ? 'row' : 'rows'}`
  const columns = `${range.width} ${range.width === 1 ? 'column' : 'columns'}`
  const groups: ReadonlyArray<ReadonlyArray<readonly [string, () => void]>> = [
    [
      [`Insert ${rows} above`, () => workbook.insertRows(sheetId, range.y, range.height)],
      [`Insert ${rows} below`, () => workbook.insertRows(sheetId, range.y + range.height, range.height)],
      [`Insert ${columns} left`, () => workbook.insertColumns(sheetId, range.x, range.width)],
      [`Insert ${columns} right`, () => workbook.insertColumns(sheetId, range.x + range.width, range.width)],
    ],
    [
      [`Delete ${rows}`, () => workbook.deleteRows(sheetId, range.y, range.height)],
      [`Delete ${columns}`, () => workbook.deleteColumns(sheetId, range.x, range.width)],
    ],
    [
      ['Sort A to Z by this column', () => sortByActiveColumn(controller, true)],
      ['Sort Z to A by this column', () => sortByActiveColumn(controller, false)],
    ],
    [
      ['Clear contents', () => workbook.clearContents(sheetId, range)],
      ['Clear formatting', () => workbook.clearFormatting(sheetId, range)],
    ],
  ]
  return (
    <ContextMenu.Root>
      <ContextMenu.Trigger className="sheet-canvas">{children}</ContextMenu.Trigger>
      <ContextMenu.Portal>
        <ContextMenu.Positioner>
          <ContextMenu.Popup className="menu-popup">
            {groups.map((group, index) => (
              <ContextMenu.Group key={group[0]?.[0]}>
                {index > 0 && <ContextMenu.Separator className="menu-separator" />}
                {group.map(([label, action]) => (
                  <ContextMenu.Item key={label} className="menu-item" onClick={action}><span>{label}</span></ContextMenu.Item>
                ))}
              </ContextMenu.Group>
            ))}
          </ContextMenu.Popup>
        </ContextMenu.Positioner>
      </ContextMenu.Portal>
    </ContextMenu.Root>
  )
}
