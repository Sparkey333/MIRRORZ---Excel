/**
 * VBA support: reading the macro project out of a workbook.
 *
 * Step one of macro support is showing people the code that is in their file.
 * That is all this package does. Running it is a separate problem with a
 * separate risk profile, and it is not solved here.
 */

export * from './compression.js';
export * from './project.js';
export * from './lexer.js';
export * from './parser.js';
export * from './compat.js';
