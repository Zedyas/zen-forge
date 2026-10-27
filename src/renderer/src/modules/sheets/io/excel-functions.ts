/**
 * Excel saves a function added after Excel 2007 with a prefix (`_xlfn.IFS`), and shows #NAME? for one
 * saved without it. The engine knows these by their plain names, so the prefix comes off as a file
 * opens and goes back on as it saves. Functions the engine does not have keep their prefix throughout.
 */

const sheetFunctions = 'ACOT ACOTH ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST COMBINA CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FLOOR.MATH FLOOR.PRECISE FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOWEEKNUM LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PDURATION PERCENTILE.EXC PERCENTILE.INC PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC RRI SEC SECH SEQUENCE SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TEXTJOIN UNICHAR UNICODE UNIQUE VAR.P VAR.S VSTACK WEIBULL.DIST WORKDAY.INTL XLOOKUP XOR Z.TEST'

/** The file prefix of each function the engine calculates; FILTER and SORT also carry the worksheet marker. */
const prefixes = new Map<string, string>([
  ...sheetFunctions.split(' ').map(name => [name, '_xlfn.'] as const),
  ['FILTER', '_xlfn._xlws.'],
  ['SORT', '_xlfn._xlws.'],
])

/** Text in quotes ("a""b") or a quoted sheet name ('Q1 ''26'), where a name is not a function call. */
const quoted = /"(?:[^"]|"")*"?|'(?:[^']|'')*'?/g
/** A function call with any prefixes, not part of a longer name. */
const call = /(?<![\p{L}\p{N}_.])((?:_xlfn\.|_xlws\.)*)([A-Za-z][A-Za-z0-9.]*)(?=\s*\()/gu

function outsideQuotes(formula: string, replace: (text: string) => string): string {
  let result = ''
  let last = 0
  for (const match of formula.matchAll(quoted)) {
    result += replace(formula.slice(last, match.index)) + match[0]
    last = match.index + match[0].length
  }
  return result + replace(formula.slice(last))
}

/** A formula from a file (without its `=`), as the engine reads it: `_xlfn.IFS(` becomes `IFS(`. */
export function engineFormula(formula: string): string {
  return outsideQuotes(formula, text => text.replace(call, (match, _prefix: string, name: string) =>
    prefixes.has(name.toUpperCase()) ? name : match))
}

/** A formula (without its `=`) as Excel reads it from a file: `IFS(` becomes `_xlfn.IFS(`. */
export function fileFormula(formula: string): string {
  return outsideQuotes(formula, text => text.replace(call, (match, _prefix: string, name: string) => {
    const prefix = prefixes.get(name.toUpperCase())
    return prefix === undefined ? match : `${prefix}${name.toUpperCase()}`
  }))
}
