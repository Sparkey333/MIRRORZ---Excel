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

import { MATH_FUNCTIONS } from './math.js';
import { TEXT_FUNCTIONS } from './text.js';
// --- category imports go here ---

/** Every implemented function, in registration order. */
export const ALL_FUNCTIONS: readonly FunctionSpec[] = [
  ...MATH_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  // --- category arrays go here ---
];

/** Old names that must resolve to their modern equivalents. */
export const FUNCTION_ALIASES: readonly (readonly [string, string])[] = [
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
