import type { FunctionListEntry } from 'hyperformula'
import type { Workbook } from './model/Workbook'

interface FunctionAutocompleteProps {
  readonly workbook: Workbook
  readonly input: string
  onChoose(name: string): void
}

/** The function name being typed at the end of a formula: after the = sign, an operator, a comma or a bracket. */
export function functionPrefix(input: string): string | undefined {
  const match = input.match(/(?:^|[=+\-*/,(&^<>\s])([A-Z][A-Z0-9.]*)$/i)
  if (!input.startsWith('=') || match?.[1] === undefined) return undefined
  // An odd number of quotes before the name means it is inside a text string, not a function call.
  const quotes = input.slice(0, input.length - match[1].length).split('"').length - 1
  return quotes % 2 === 0 ? match[1] : undefined
}

export function FunctionAutocomplete({ workbook, input, onChoose }: FunctionAutocompleteProps) {
  const prefix = functionPrefix(input)
  if (prefix === undefined) return null

  const suggestions = workbook.functionSuggestions(prefix)
  if (suggestions.length === 0) return null

  return (
    <div className="function-autocomplete" role="listbox" aria-label="Formula functions">
      {suggestions.map(suggestion => (
        <FunctionSuggestion
          key={suggestion.localizedName}
          workbook={workbook}
          suggestion={suggestion}
          onChoose={onChoose}
        />
      ))}
    </div>
  )
}

interface FunctionSuggestionProps {
  readonly workbook: Workbook
  readonly suggestion: FunctionListEntry
  onChoose(name: string): void
}

function FunctionSuggestion({ workbook, suggestion, onChoose }: FunctionSuggestionProps) {
  const details = workbook.functionDetails(suggestion.canonicalName)
  const signature = details === undefined
    ? `${suggestion.localizedName}(…)`
    : `${suggestion.localizedName}(${details.parameters.map(parameter =>
        parameter.optional ? `[${parameter.name}]` : parameter.name,
      ).join(', ')}${details.repeatLastArgs > 0 ? ', …' : ''})`

  return (
    <button
      type="button"
      role="option"
      aria-selected="false"
      onMouseDown={event => event.preventDefault()}
      onClick={() => onChoose(suggestion.localizedName)}
    >
      <span>
        <strong>{suggestion.localizedName}</strong>
        <code>{signature}</code>
      </span>
      <small>{suggestion.shortDescription ?? suggestion.category}</small>
    </button>
  )
}
