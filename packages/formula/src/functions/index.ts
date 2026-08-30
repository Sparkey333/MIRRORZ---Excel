/**
 * The worksheet function library.
 *
 * Functions are grouped by Excel's own categories, one module each, and every
 * module exports a flat array of FunctionSpec. Splitting them this way keeps
 * each file reviewable and lets categories be built and tested independently.
 *
 * To add a category: import its array below and spread it into ALL_FUNCTIONS.
 * Compatibility aliases - pre-2010 names that old files still use, such as
 * NORMDIST for NORM.DIST - go in FUNCTION_ALIASES so a workbook written in 2008
 * does not show #NAME?.
 */

import { FunctionRegistry, type FunctionSpec } from '../registry.js';

import { DATETIME_FUNCTIONS } from './datetime.js';
import { DYNAMIC_ARRAY_FUNCTIONS } from './dynamic.js';
import { ENGINEERING_FUNCTIONS } from './engineering.js';
import { FINANCIAL_FUNCTIONS } from './financial.js';
import { LOGICAL_FUNCTIONS } from './logical.js';
import { LOOKUP_FUNCTIONS } from './lookup.js';
import { MATH_FUNCTIONS } from './math.js';
import { STATISTICAL_FUNCTIONS } from './statistical.js';
import { TEXT_FUNCTIONS } from './text.js';
// --- category imports go here ---

/** Every implemented function, in registration order. */
export const ALL_FUNCTIONS: readonly FunctionSpec[] = [
  ...DATETIME_FUNCTIONS,
  ...DYNAMIC_ARRAY_FUNCTIONS,
  ...ENGINEERING_FUNCTIONS,
  ...FINANCIAL_FUNCTIONS,
  ...LOGICAL_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...MATH_FUNCTIONS,
  ...STATISTICAL_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  // --- category arrays go here ---
];

/** Old names that must resolve to their modern equivalents. */
export const FUNCTION_ALIASES: readonly (readonly [string, string])[] = [
  // Statistical: the pre-2010 spellings, which every file written before Excel
  // 2010 still carries.
  ['NORMDIST', 'NORM.DIST'],
  ['NORMINV', 'NORM.INV'],
  ['NORMSDIST', 'NORM.S.DIST'],
  ['NORMSINV', 'NORM.S.INV'],
  ['TDIST', 'T.DIST.2T'],
  ['TINV', 'T.INV.2T'],
  ['FDIST', 'F.DIST.RT'],
  ['FINV', 'F.INV.RT'],
  ['CHIDIST', 'CHISQ.DIST.RT'],
  ['CHIINV', 'CHISQ.INV.RT'],
  ['BINOMDIST', 'BINOM.DIST'],
  ['NEGBINOMDIST', 'NEGBINOM.DIST'],
  ['POISSON', 'POISSON.DIST'],
  ['HYPGEOMDIST', 'HYPGEOM.DIST'],
  ['EXPONDIST', 'EXPON.DIST'],
  ['WEIBULL', 'WEIBULL.DIST'],
  ['LOGNORMDIST', 'LOGNORM.DIST'],
  ['LOGINV', 'LOGNORM.INV'],
  ['BETADIST', 'BETA.DIST'],
  ['BETAINV', 'BETA.INV'],
  ['GAMMADIST', 'GAMMA.DIST'],
  ['GAMMAINV', 'GAMMA.INV'],
  ['CONFIDENCE', 'CONFIDENCE.NORM'],
  ['ZTEST', 'Z.TEST'],
  ['TTEST', 'T.TEST'],
  ['FTEST', 'F.TEST'],
  ['CHITEST', 'CHISQ.TEST'],
  ['COVAR', 'COVARIANCE.P'],
  ['STDEV', 'STDEV.S'],
  ['STDEVP', 'STDEV.P'],
  ['VAR', 'VAR.S'],
  ['VARP', 'VAR.P'],
  ['PERCENTILE', 'PERCENTILE.INC'],
  ['QUARTILE', 'QUARTILE.INC'],
  ['PERCENTRANK', 'PERCENTRANK.INC'],
  ['MODE', 'MODE.SNGL'],
  ['RANK', 'RANK.EQ'],
  ['CRITBINOM', 'BINOM.INV'],
  // --- aliases go here ---
];

export function createRegistry(): FunctionRegistry {
  const registry = new FunctionRegistry();
  registry.registerAll(ALL_FUNCTIONS);
  for (const [oldName, canonical] of FUNCTION_ALIASES) {
    // Skip an alias whose target has not been implemented yet, so a partially
    // populated library still builds.
    if (registry.has(canonical) && !registry.has(oldName)) registry.alias(oldName, canonical);
  }
  return registry;
}
