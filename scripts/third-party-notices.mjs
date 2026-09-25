// Collects the licenses a packaged Zendo has to ship: Zendo's own, Electron's and Chromium's, and the
// license text of every production dependency. Writes them to out/licenses/, which electron-builder
// copies into Zendo.app/Contents/Resources/licenses/. Run by `pnpm release:mac` and `pnpm package:mac`.
import { execFileSync } from 'node:child_process'
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join } from 'node:path'

const require = createRequire(import.meta.url)
const outDir = 'out/licenses'
mkdirSync(outDir, { recursive: true })

// Requiring electron downloads its binary if it is missing; its dist folder holds Chromium's license list.
const electronDist = join(require('electron'), '../../../..')
copyFileSync('LICENSE', join(outDir, 'LICENSE'))
copyFileSync(join(electronDist, 'LICENSE'), join(outDir, 'LICENSE.electron.txt'))
copyFileSync(join(electronDist, 'LICENSES.chromium.html'), join(outDir, 'LICENSES.chromium.html'))

const byLicense = JSON.parse(execFileSync('pnpm', ['licenses', 'list', '--prod', '--json'], { encoding: 'utf8' }))
const packages = Object.values(byLicense).flat().sort((a, b) => a.name.localeCompare(b.name))

function licenseText(directory) {
  const file = existsSync(directory)
    ? readdirSync(directory).find(name => /^(licen[cs]e|copying|notice)/i.test(name))
    : undefined
  return file === undefined ? undefined : readFileSync(join(directory, file), 'utf8').trim()
}

const sections = packages.map(({ name, versions, license, homepage, paths }) => {
  const text = paths.map(licenseText).find(value => value !== undefined)
  const heading = `${name} ${versions.join(', ')} (${license})`
  return [heading, '-'.repeat(heading.length), text ?? `No license file is included in this package. See ${homepage ?? 'its npm page'}.`].join('\n')
})

writeFileSync(join(outDir, 'THIRD-PARTY-NOTICES.txt'), [
  'Third-party software in Zendo',
  '',
  'Zendo includes the open-source packages below, each under its own license. Electron and Chromium',
  'licenses are in LICENSE.electron.txt and LICENSES.chromium.html in this folder.',
  '',
  ...sections.map(section => `${section}\n`),
].join('\n'))
console.log(`Licenses written to ${outDir}/ (${packages.length} packages).`)
