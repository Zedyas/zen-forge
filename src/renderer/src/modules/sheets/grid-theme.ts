import type { Theme } from '@glideapps/glide-data-grid'

export const rowHeight = 26
export const headerHeight = 26
export const cellFontSize = 13

function rgb(hex: string): readonly [number, number, number] {
  const value = Number.parseInt(hex.replace('#', ''), 16)
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255]
}

/** `hex` at `alpha`, as rgba(). Glide blends theme colours with its own parser, which rejects color-mix(). */
function withAlpha(hex: string, alpha: number): string {
  const [red, green, blue] = rgb(hex)
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`
}

/** `top` laid over `bottom` at `amount`, as an opaque hex colour. */
function mix(top: string, bottom: string, amount: number): string {
  const upper = rgb(top)
  const lower = rgb(bottom)
  return `#${upper.map((channel, index) => Math.round(channel * amount + (lower[index] ?? 0) * (1 - amount)).toString(16).padStart(2, '0')).join('')}`
}

/** Maps the suite's CSS tokens onto Glide's theme; recomputed when the appearance changes. */
export function gridTheme(): Partial<Theme> {
  const styles = getComputedStyle(document.documentElement)
  const token = (name: string): string => styles.getPropertyValue(name).trim()
  const accent = token('--accent')
  return {
    accentColor: accent,
    accentLight: withAlpha(accent, 0.13),
    accentFg: token('--on-accent'),
    textDark: token('--ink'),
    textMedium: token('--dim'),
    textLight: token('--faint'),
    textHeader: token('--dim'),
    textHeaderSelected: token('--ink'),
    textBubble: token('--ink'),
    bgIconHeader: token('--dim'),
    fgIconHeader: token('--doc'),
    bgCell: token('--doc'),
    bgCellMedium: token('--raised'),
    bgHeader: token('--raised'),
    bgHeaderHovered: token('--sunken'),
    bgHeaderHasFocus: mix(accent, token('--raised'), 0.16),
    bgBubble: token('--raised'),
    bgSearchResult: withAlpha(token('--warn'), 0.25),
    borderColor: token('--grid-line'),
    horizontalBorderColor: token('--grid-line'),
    headerBottomBorderColor: token('--line'),
    drilldownBorder: token('--line'),
    linkColor: accent,
    fontFamily: token('--font-ui'),
    baseFontStyle: `${cellFontSize}px`,
    headerFontStyle: '500 11.5px',
    markerFontStyle: '11px',
    editorFontSize: `${cellFontSize}px`,
    cellHorizontalPadding: 8,
    roundingRadius: 3,
  }
}

/** Relative luminance check so text stays readable on a light fill in dark mode. */
export function isLightColor(hex: string): boolean {
  const [red, green, blue] = rgb(hex)
  return (0.299 * red + 0.587 * green + 0.114 * blue) / 255 > 0.6
}
