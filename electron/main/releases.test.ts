import { describe, expect, it } from 'vitest'
import { checksumFor, compareVersions, parseVersion, selectRelease } from './releases'

function order(a: string, b: string): number {
  const left = parseVersion(a)
  const right = parseVersion(b)
  if (left === undefined || right === undefined) throw new Error(`Not versions: ${a}, ${b}`)
  return Math.sign(compareVersions(left, right))
}

describe('compareVersions', () => {
  it('orders by number, not by text', () => {
    expect(order('0.10.0', '0.9.9')).toBe(1)
    expect(order('v1.2.3', '1.2.3')).toBe(0)
    expect(order('1.0.0', '2.0.0')).toBe(-1)
  })

  it('puts a pre-release before its release, in semver order', () => {
    const sorted = ['1.0.0', '1.0.0-rc.1', '1.0.0-beta.11', '1.0.0-alpha', '1.0.0-beta.2', '1.0.0-alpha.1', '1.0.0-beta']
      .sort((a, b) => order(a, b))
    expect(sorted).toEqual(['1.0.0-alpha', '1.0.0-alpha.1', '1.0.0-beta', '1.0.0-beta.2', '1.0.0-beta.11', '1.0.0-rc.1', '1.0.0'])
  })

  it('rejects text that is not a version', () => {
    expect(parseVersion('latest')).toBeUndefined()
    expect(parseVersion('1.2')).toBeUndefined()
  })
})

const download = 'https://github.com/Zedyas/zen-forge/releases/download'

function release(tag: string, options: { draft?: boolean; prerelease?: boolean; assets?: readonly string[] } = {}) {
  const version = tag.replace(/^v/, '')
  const names = options.assets ?? [`Zendo-${version}-arm64.dmg`, 'SHA256SUMS.txt']
  return {
    tag_name: tag,
    draft: options.draft ?? false,
    prerelease: options.prerelease ?? true,
    html_url: `https://github.com/Zedyas/zen-forge/releases/tag/${tag}`,
    assets: names.map(name => ({ name, size: 100, browser_download_url: `${download}/${tag}/${name}` })),
  }
}

describe('selectRelease', () => {
  it('picks the highest newer release that has both files, skipping drafts', () => {
    const body = [
      release('v0.4.0', { draft: true }),
      release('v0.3.0', { assets: ['Zendo-0.3.0-arm64.dmg'] }),
      release('v0.2.0'),
      release('v0.2.1'),
      release('v0.1.1'),
    ]
    expect(selectRelease(body, '0.1.1', 'arm64')).toEqual({
      version: '0.2.1',
      pageUrl: 'https://github.com/Zedyas/zen-forge/releases/tag/v0.2.1',
      dmg: { name: 'Zendo-0.2.1-arm64.dmg', url: `${download}/v0.2.1/Zendo-0.2.1-arm64.dmg`, size: 100 },
      checksumsUrl: `${download}/v0.2.1/SHA256SUMS.txt`,
    })
  })

  it('offers nothing when the newest release is the running version', () => {
    expect(selectRelease([release('v0.1.1'), release('v0.1.0')], '0.1.1', 'arm64')).toBeUndefined()
  })

  it('ignores pre-releases once a full release exists', () => {
    const body = [release('v0.3.0', { prerelease: true }), release('v0.2.0', { prerelease: false })]
    expect(selectRelease(body, '0.1.1', 'arm64')?.version).toBe('0.2.0')
  })

  it('skips a dmg for another architecture or hosted outside the releases', () => {
    const elsewhere = { ...release('v0.2.0'), assets: release('v0.2.0').assets.map(asset => ({ ...asset, browser_download_url: `${download}/../../../other/repo/${asset.name}` })) }
    expect(selectRelease([release('v0.2.0')], '0.1.1', 'x64')).toBeUndefined()
    expect(selectRelease([elsewhere], '0.1.1', 'arm64')).toBeUndefined()
  })
})

describe('checksumFor', () => {
  const hash = 'A'.repeat(64)

  it('reads the line for the file, in text or binary mode', () => {
    expect(checksumFor(`${'b'.repeat(64)}  Other.dmg\n${hash}  Zendo-0.2.0-arm64.dmg\n`, 'Zendo-0.2.0-arm64.dmg')).toBe('a'.repeat(64))
    expect(checksumFor(`${hash} *Zendo-0.2.0-arm64.dmg\r\n`, 'Zendo-0.2.0-arm64.dmg')).toBe('a'.repeat(64))
  })

  it('finds nothing for a missing file or a malformed line', () => {
    expect(checksumFor(`${hash}  Zendo-0.2.0-arm64.dmg.part`, 'Zendo-0.2.0-arm64.dmg')).toBeUndefined()
    expect(checksumFor(`${'a'.repeat(63)}  Zendo-0.2.0-arm64.dmg`, 'Zendo-0.2.0-arm64.dmg')).toBeUndefined()
  })
})
