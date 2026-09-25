import { toast } from 'sonner'
import type { Paper } from '@shared/print'
import { platformClient } from '../platform/client'

/*
 * Printing for every module. A module builds a printout (content plus page setup) for the paper
 * it is given. This service mounts it into a print-only container (styles.css hides the app and
 * shows only that container when printing), waits for its images, prints the window's page, and
 * removes the printout again.
 */

export type Orientation = 'portrait' | 'landscape'

/** A sheet of paper and the blank space kept on every side, in millimetres. */
export interface PageSetup {
  readonly width: number
  readonly height: number
  readonly margin: number
}

export interface Printout {
  /** What prints. It is mounted only while the job runs, so it never shows on screen. */
  readonly content: Node
  /** The page setup for the content. */
  readonly page: PageSetup
  /**
   * More page setups by name, for printouts whose pages differ (a PDF mixing orientations). An
   * element styled `page: <name>` starts a new sheet with that setup (CSS named pages).
   */
  readonly namedPages?: Readonly<Record<string, PageSetup>>
  /** Frees what the content holds, such as blob URLs. Runs once the job has ended. */
  readonly release?: () => void
}

/** Builds a printout for the paper, or resolves undefined when there is nothing to print. */
export type PrintoutBuilder = (paper: Paper) => Printout | undefined | Promise<Printout | undefined>

/** `paper` turned to `orientation`, with `margin` millimetres on every side. */
export function pageSetup(paper: Paper, orientation: Orientation, margin: number): PageSetup {
  return orientation === 'portrait'
    ? { width: paper.width, height: paper.height, margin }
    : { width: paper.height, height: paper.width, margin }
}

/** The area inside the margins, in millimetres. */
export function printableArea(page: PageSetup): { readonly width: number; readonly height: number } {
  return { width: page.width - 2 * page.margin, height: page.height - 2 * page.margin }
}

/** The white background covers the margins, which Chromium otherwise fills with the window's own background colour. */
function pageRule(name: string, page: PageSetup): string {
  return `@page ${name} { size: ${page.width}mm ${page.height}mm; margin: ${page.margin}mm; background: #fff; }`
}

/** An image that has not loaded when printing starts prints as a blank. */
function imagesLoaded(root: HTMLElement): Promise<unknown> {
  return Promise.all(Array.from(root.querySelectorAll('img'), image => new Promise((resolve, reject) => {
    const failed = (): void => reject(new Error('Part of the printout could not be drawn.'))
    if (image.complete) {
      if (image.naturalWidth > 0) resolve(undefined)
      else failed()
      return
    }
    image.addEventListener('load', resolve, { once: true })
    image.addEventListener('error', failed, { once: true })
  })))
}

let busy = false

/** Builds and mounts a printout, runs `job` on the page, then removes it. A request while one runs is ignored. */
async function withPrintout<Result>(build: PrintoutBuilder, job: () => Promise<Result>): Promise<Result | undefined> {
  if (busy) {
    toast.info('Zendo is already preparing a printout', { description: 'Try again when it has finished.' })
    return undefined
  }
  busy = true
  const root = document.createElement('div')
  let printout: Printout | undefined
  try {
    printout = await build(await platformClient.printPaper())
    if (printout === undefined) return undefined
    const style = document.createElement('style')
    const named = Object.entries(printout.namedPages ?? {}).map(([name, page]) => pageRule(name, page))
    style.textContent = [pageRule('', printout.page), ...named].join('\n')
    root.className = 'print-root'
    root.append(style, printout.content)
    document.body.append(root)
    await imagesLoaded(root)
    return await job()
  } finally {
    root.remove()
    printout?.release?.()
    busy = false
  }
}

/** Opens the print dialog for a printout. Resolves true once printed; false when cancelled, when there was nothing to print, or while another printout is being prepared. */
export async function openPrintDialog(build: PrintoutBuilder): Promise<boolean> {
  return (await withPrintout(build, () => platformClient.print())) ?? false
}

/** Renders a printout to PDF bytes; undefined when there was nothing to print or another printout is being prepared. */
export function printToPdf(build: PrintoutBuilder): Promise<Uint8Array | undefined> {
  return withPrintout(build, () => platformClient.printToPdf())
}
