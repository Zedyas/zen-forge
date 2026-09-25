import { useEffect, useRef, useState, type FormEvent } from 'react'
import type { Item } from '@glideapps/glide-data-grid'
import type { Workbook } from './model/Workbook'
import { cellName, formulaReferences } from './model/address'
import { FunctionAutocomplete } from './FunctionAutocomplete'

interface FormulaBarProps {
  readonly workbook: Workbook
  readonly sheetId: number
  readonly selectedCell: Item
}

export function FormulaBar({ workbook, sheetId, selectedCell }: FormulaBarProps) {
  const currentInput = workbook.getCellInput(sheetId, selectedCell)
  const [draft, setDraft] = useState(currentInput)
  // The draft resets only when the selected cell moves or its stored input changes. The parent
  // rebuilds `selectedCell` on every render (a save, a toolbar toggle), so compare its address,
  // not the array, or in-progress typing is wiped.
  const address = cellName(selectedCell)
  const committedRef = useRef(currentInput)
  const previousAddressRef = useRef(address)

  useEffect(() => {
    const cellChanged = previousAddressRef.current !== address
    previousAddressRef.current = address
    if (cellChanged || currentInput !== committedRef.current) {
      committedRef.current = currentInput
      setDraft(currentInput)
    }
  }, [currentInput, address])

  const commit = (event?: FormEvent): void => {
    event?.preventDefault()
    committedRef.current = draft
    workbook.setCellInput(sheetId, selectedCell, draft)
  }

  const chooseFunction = (name: string): void => {
    const prefixMatch = draft.match(/[A-Z][A-Z0-9.]*$/i)
    if (prefixMatch === null) return
    setDraft(`${draft.slice(0, prefixMatch.index)}${name}(`)
  }

  const references = formulaReferences(draft)

  return (
    <form className="formula-bar" onSubmit={commit}>
      <output className="formula-address" aria-label="Selected cell">{cellName(selectedCell)}</output>
      <span className="formula-fx" aria-hidden="true">fx</span>
      <div className="formula-editor">
        <input
          value={draft}
          onChange={event => setDraft(event.currentTarget.value)}
          onBlur={() => commit()}
          aria-label="Cell value or formula"
          spellCheck="false"
          autoComplete="off"
        />
        <FunctionAutocomplete workbook={workbook} input={draft} onChoose={chooseFunction} />
      </div>
      {references.length > 0 && (
        <div className="formula-references" aria-label="Formula references">
          {references.map((reference, index) => (
            <span key={reference} data-reference-color={index % 4}>{reference}</span>
          ))}
        </div>
      )}
    </form>
  )
}
