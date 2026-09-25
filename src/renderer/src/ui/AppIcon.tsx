import type { ApplicationId } from '@shared/applications'
import { ensoPath, stoneRadius, stones } from './suite-mark'

interface AppIconProps {
  readonly application: ApplicationId
  readonly size?: number
  readonly unavailable?: boolean
}

/*
 * Custom artwork: no approved icon set has a PDF with a hanko seal or an ensō. Plates use the same
 * four colours as the suite mark's stones and each tab's accent.
 * Every application icon is a white paper object on a plate of its colour, so the
 * paper says what the file is (grid, "PDF", text lines, slide) and one Japanese
 * detail carries the name: the vermilion seal on Hanko, the ink brush stroke on Sumi.
 * The plates keep their colour in dark mode because the icons are artwork, like Dock icons.
 */

const plate = { x: 1, y: 1, width: 30, height: 30, rx: 8 }

function SheetsArt() {
  return (
    <>
      <rect {...plate} fill="#2fb592" />
      <rect x="7.5" y="6.5" width="17" height="19" rx="2" fill="#fff" />
      <path d="M7.5 8.5a2 2 0 0 1 2-2h13a2 2 0 0 1 2 2v2.5h-17z" fill="#cfeee4" />
      <path d="M7.5 15.5h17M7.5 20.5h17M13 11v14.5" stroke="#2fb592" strokeOpacity=".7" strokeWidth="1.1" />
    </>
  )
}

function PdfArt() {
  return (
    <>
      <rect {...plate} fill="#f06a52" />
      <path d="M9.5 5.5h9l5 5v15a1.5 1.5 0 0 1-1.5 1.5h-12.5a1.5 1.5 0 0 1-1.5-1.5v-18.5a1.5 1.5 0 0 1 1.5-1.5z" fill="#fff" />
      <path d="M18.5 5.5v3.5a1.5 1.5 0 0 0 1.5 1.5h3.5" fill="#f3d3cc" />
      <text x="9.6" y="16.4" fill="#23252a" fontFamily="system-ui, -apple-system, sans-serif" fontSize="6.4" fontWeight="800" letterSpacing="-0.2">PDF</text>
      <circle cx="19.2" cy="21.6" r="3.9" fill="#f06a52" />
      <circle cx="19.2" cy="21.6" r="2.9" fill="none" stroke="#fff" strokeWidth=".7" />
      <path d="M18.1 20.5h2.2M19.2 20.5v2.4M18.1 22.9h2.2" stroke="#fff" strokeWidth=".75" strokeLinecap="round" />
    </>
  )
}

function DocsArt() {
  return (
    <>
      <rect {...plate} fill="#7d97df" />
      <rect x="8" y="5.5" width="16" height="21" rx="1.5" fill="#fff" />
      <path d="M11 9.5h7M11 12.5h10M11 15.5h10M11 18.5h6" stroke="#b8c2d6" strokeWidth="1.3" strokeLinecap="round" />
      <path d="M10.2 23.4c2.6-1.9 5.4-2.6 8.1-2.4 1.5.1 2.8.5 4.1 1.1-1.3-.3-2.7-.3-4-.1-2.8.4-5.4 1.2-8.2 1.4z" fill="#1c1e24" />
    </>
  )
}

function SlidesArt() {
  return (
    <>
      <rect {...plate} fill="#e0a92c" />
      <rect x="5.5" y="7.5" width="21" height="14" rx="1.5" fill="#fff" />
      <path d="M8.5 11h8" stroke="#23252a" strokeWidth="1.4" strokeLinecap="round" />
      <path d="M9 18.5v-3M12.5 18.5v-5M16 18.5v-2" stroke="#e0a92c" strokeWidth="2" strokeLinecap="round" />
      <path d="M16 21.5v3.5M12.5 25.5h7" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" />
    </>
  )
}

/**
 * The suite mark for the Home tab: the ensō and its four stones without the Dock icon's plate, so it
 * sits in the title bar like the other tab icons instead of a box inside the tab's box. The ensō
 * takes the text colour, so it reads in light and dark. (The Dock icon comes from build/icon.png.)
 */
function SuiteArt() {
  return (
    <>
      <path fill="currentColor" d={ensoPath} />
      {stones.map(stone => (
        <ellipse key={stone.colour} cx={stone.x} cy={stone.y} rx={stoneRadius + 0.2} ry={stoneRadius} fill={stone.colour} />
      ))}
    </>
  )
}

const artwork: Record<ApplicationId, () => React.JSX.Element> = {
  home: SuiteArt,
  sheets: SheetsArt,
  pdf: PdfArt,
  docs: DocsArt,
  slides: SlidesArt,
}

/** The suite mark has no plate, so its view is cropped to the ensō and it fills the size it is given. */
const viewBoxes: Partial<Record<ApplicationId, string>> = { home: '4.5 4.2 21.7 22.4' }

export function AppIcon({ application, size = 20, unavailable = false }: AppIconProps) {
  const Art = artwork[application]
  return (
    <svg
      className={`app-icon${unavailable ? ' is-unavailable' : ''}`}
      width={size}
      height={size}
      viewBox={viewBoxes[application] ?? '0 0 32 32'}
      aria-hidden="true"
    >
      <Art />
    </svg>
  )
}
