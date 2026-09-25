/*
 * Props for Lucide icons, spread onto an icon (`<Undo2 {...icon} />`) so icons in one place share a size
 * and a line weight. The sizes are --icon and --icon-small in styles.css; the stroke widths give each
 * size a line of about 1px.
 */

/** Bars (title bar, toolbar, find bar), menus and Home. */
export const icon = { size: 16, strokeWidth: 1.7, 'aria-hidden': true } as const

/** The status bar, inspector, panels, popovers and toasts, and the check mark in menus. */
export const smallIcon = { size: 14, strokeWidth: 2, 'aria-hidden': true } as const

/** Close marks on tabs and the carets on menu buttons. */
export const tinyIcon = { size: 12, strokeWidth: 2, 'aria-hidden': true } as const
