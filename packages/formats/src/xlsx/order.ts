/**
 * Child-element ordering for the OOXML complex types we serialise.
 *
 * OOXML schemas use xs:sequence, not xs:all. An element that is otherwise
 * perfectly valid but sits in the wrong position is the single most common cause
 * of Excel's "we found a problem with some content" repair dialog - and the
 * recovery log names only the part, never the element, which makes it a
 * miserable bug to find after the fact.
 *
 * Keeping the order as data rather than as the shape of a render function means
 * the serialiser is physically incapable of appending out of order: it iterates
 * this table and emits whatever fragment it holds for each name. Hand-maintained
 * emit sequences drift, and at least one widely used library currently ships a
 * worksheet writer that emits legacyDrawing after extLst rather than at
 * position 31, which is exactly the defect this design prevents.
 *
 * Positions follow ECMA-376 Part 1, CT_Worksheet, CT_Workbook and CT_Stylesheet.
 */

/** CT_Worksheet child order. */
export const WORKSHEET_CHILD_ORDER = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'drawingHF',
  // legacyDrawing belongs immediately after the drawing elements, well before
  // picture, tableParts and extLst.
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
] as const;

/** CT_Workbook child order. */
export const WORKBOOK_CHILD_ORDER = [
  'fileVersion',
  'fileSharing',
  'workbookPr',
  'workbookProtection',
  'bookViews',
  'sheets',
  'functionGroups',
  'externalReferences',
  'definedNames',
  'calcPr',
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
] as const;

/** CT_Stylesheet child order. */
export const STYLESHEET_CHILD_ORDER = [
  'numFmts',
  'fonts',
  'fills',
  'borders',
  'cellStyleXfs',
  'cellXfs',
  'cellStyles',
  'dxfs',
  'tableStyles',
  'colors',
  'extLst',
] as const;

export type WorksheetChild = (typeof WORKSHEET_CHILD_ORDER)[number];

/**
 * Sort raw XML fragments into schema order.
 *
 * Fragments are keyed by element local name. Anything not in the order table is
 * emitted last, which is the least-bad answer for an element from a schema
 * extension we do not know about: dropping it would lose data, and guessing a
 * position could invalidate the sequence.
 */
export function emitInOrder(
  order: readonly string[],
  fragments: ReadonlyMap<string, string>,
): string {
  const parts: string[] = [];
  const seen = new Set<string>();
  for (const name of order) {
    const fragment = fragments.get(name);
    if (fragment !== undefined) {
      parts.push(fragment);
      seen.add(name);
    }
  }
  for (const [name, fragment] of fragments) {
    if (!seen.has(name)) parts.push(fragment);
  }
  return parts.join('');
}

/** Position of an element in its parent's sequence, or -1 when unknown. */
export function orderIndex(order: readonly string[], name: string): number {
  return order.indexOf(name);
}
