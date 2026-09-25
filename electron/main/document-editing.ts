import { Menu, shell, type MenuItemConstructorOptions, type WebContents } from 'electron'
import { handle } from './security'

/** Links in documents open in the default browser, never in Zendo. Only web links: no files, scripts or other apps. */
export function registerLinkIpc(): void {
  handle('shell:open-external', async (_event, url: string) => {
    const parsed = new URL(url)
    if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') throw new Error('Only web links can be opened.')
    await shell.openExternal(parsed.toString())
  })
}

/**
 * Right-clicking a misspelled word in editable text offers macOS's corrections. Every other
 * right-click is left to the page, and Zendo's own context menus cancel the event, so this
 * never shows over them.
 */
export function attachSpellingMenu(contents: WebContents): void {
  contents.on('context-menu', (_event, params) => {
    if (!params.isEditable || params.misspelledWord === '') return
    const suggestions: MenuItemConstructorOptions[] = params.dictionarySuggestions.slice(0, 6)
      .map(suggestion => ({ label: suggestion, click: () => contents.replaceMisspelling(suggestion) }))
    Menu.buildFromTemplate([
      ...(suggestions.length > 0 ? suggestions : [{ label: 'No Guesses Found', enabled: false }]),
      { type: 'separator' },
      { label: 'Learn Spelling', click: () => contents.session.addWordToSpellCheckerDictionary(params.misspelledWord) },
      { type: 'separator' },
      { role: 'cut' },
      { role: 'copy' },
      { role: 'paste' },
    ]).popup()
  })
}
