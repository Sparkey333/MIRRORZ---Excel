/**
 * Engineering and database functions.
 *
 * Six decisions shape this file.
 *
 * First, the base-conversion family is two's-complement over a fixed width, not
 * a general radix conversion: ten binary digits, thirty octal bits, forty hex
 * bits. DEC2BIN(-1) is 1111111111 because the tenth bit is a sign bit, and the
 * legal input range follows from the width rather than from anything about the
 * radix. Everything - the range check, the negative encoding, the `places`
 * padding - is therefore expressed once in terms of a bit width, and each
 * function only says how wide it is.
 *
 * Second, the IM family is text in and text out, and the suffix is data. A
 * complex number written 3+4j must come back with a j, so the parsed form
 * carries the suffix it was written with and every operation propagates it;
 * mixing i and j in one call is #VALUE!, not a silent choice of one. Parsing is
 * deliberately strict - Excel accepts only a lowercase i or j and only the
 * x+yi shape - because a loose parser would turn a typo into a plausible wrong
 * number instead of the #NUM! Excel reports.
 *
 * Third, the special functions are computed from integral representations
 * rather than from the Abramowitz-and-Stegun polynomial fits that Excel itself
 * uses. J and I come from their Fourier-integral forms, where the trapezoidal
 * rule on a periodic analytic integrand converges geometrically and the aliasing
 * error is a Bessel function of very high order - which is to say zero in double
 * precision. Y and K come from their Schlaefli integrals under composite
 * Gauss-Legendre. The result is accurate to near machine precision, which is
 * around eight digits better than Excel's own BESSEL functions; since we cannot
 * be bit-identical to Excel here either way, being right is the better failure
 * mode. The one place a series is used instead is where the integral would be
 * catastrophically cancelled: J_n(x) for x small beside n is around 1e-20 while
 * the integrand is order one, so the ascending series - whose terms decay from
 * the first one in exactly that regime - takes over.
 *
 * Fourth, CONVERT is a table, and the table is the interesting part. Every unit
 * reduces to a base unit of its group by an affine map, not a ratio, because
 * temperature has offsets; the ratio-only units simply have a zero offset. A
 * metric prefix multiplies the value in its own unit, and is raised to the
 * unit's dimension so that a prefix on m^2 scales by its square. Units are
 * matched exactly before any prefix is tried, so `kn` stays a knot rather than
 * becoming a kilonewton.
 *
 * Fifth, the D functions read their criteria the way Excel's advanced filter
 * does, which is not the way COUNTIF does: rows are OR-ed, columns within a row
 * are AND-ed, a blank criteria cell constrains nothing, and a bare text
 * criterion matches any value that *begins with* it. That last rule is the one
 * that surprises people, and it is the documented behaviour of every criteria
 * range in Excel.
 *
 * Sixth, engineering functions reject logical values. TRUE is 1 everywhere else
 * in the library, but DEC2BIN(TRUE) and IMREAL(TRUE) are #VALUE! in Excel, and
 * the distinction is worth keeping because it is a real difference in the
 * arguments these functions will accept from a user.
 */

import { CellError, type Scalar, isError } from '@mirrorz/core';
import type { FunctionContext, FunctionSpec } from '../registry.js';
import { p } from '../registry.js';
import {
  type ArrayValue,
  type Criterion,
  type Value,
  arrayAt,
  checkMagnitude,
  excelAdd,
  formatNumberForConcat,
  isArray,
  isRef,
  makeArray,
  matchesCriterion,
  parseCriterion,
  toNumber,
  toText,
} from '../value.js';

// ---------------------------------------------------------------------------
// Argument plumbing
// ---------------------------------------------------------------------------

/** The scalar an ArgKind.Scalar parameter delivered, with omissions as blank. */
function scalarArg(v: Value | undefined): Scalar {
  if (v === undefined) return null;
  if (isArray(v)) return v.data[0] ?? null;
  if (isRef(v)) return null;
  return v;
}

/**
 * A number for an engineering function: numeric text is accepted, TRUE and
 * FALSE are not.
 */
function engNumber(v: Value | undefined): number | CellError {
  const s = scalarArg(v);
  if (typeof s === 'boolean') return CellError.VALUE;
  return toNumber(s);
}

/** The same, truncated towards zero the way Excel truncates these arguments. */
function engInteger(v: Value | undefined): number | CellError {
  const n = engNumber(v);
  return isError(n) ? n : Math.trunc(n);
}

/** Every scalar inside a repeating argument, flattened in row-major order. */
function* flatten(args: readonly (Value | undefined)[]): Generator<Scalar> {
  for (const arg of args) {
    if (arg === undefined) continue;
    if (isArray(arg)) {
      for (const cell of arg.data) yield cell;
      continue;
    }
    if (isRef(arg)) continue;
    yield arg;
  }
}

// ---------------------------------------------------------------------------
// Base conversion
// ---------------------------------------------------------------------------

/**
 * Widths of the two's-complement representations, in bits. Ten digits in each
 * radix: ten bits of binary, thirty of octal, forty of hexadecimal.
 */
const BIN_BITS = 10;
const OCT_BITS = 30;
const HEX_BITS = 40;

const DIGITS = '0123456789ABCDEF';

/** The text of a base-conversion input, which may have been typed as a number. */
function numeralArg(v: Value | undefined): string | CellError {
  const s = scalarArg(v);
  if (typeof s === 'boolean') return CellError.VALUE;
  if (isError(s)) return s;
  // A blank cell is zero, as it is in arithmetic; an empty string is not a
  // numeral and fails the digit check below.
  if (s === null) return '0';
  if (typeof s === 'number') {
    if (!Number.isInteger(s)) return CellError.NUM;
    return formatNumberForConcat(s);
  }
  return s.trim();
}

/** Decode ten-digit two's complement in the given radix. */
function fromNumeral(text: string, radix: number, bits: number): number | CellError {
  if (text.length === 0 || text.length > 10) return CellError.NUM;
  let magnitude = 0;
  for (const ch of text) {
    const digit = DIGITS.indexOf(ch.toUpperCase());
    if (digit < 0 || digit >= radix) return CellError.NUM;
    magnitude = magnitude * radix + digit;
  }
  const span = 2 ** bits;
  return magnitude >= span / 2 ? magnitude - span : magnitude;
}

/**
 * Encode as ten-digit two's complement, padded to `places`.
 *
 * A negative value fills the whole width, and Excel documents `places` as
 * ignored in that case - not merely redundant: DEC2HEX(-54, 6) is the ten
 * character FFFFFFFFCA rather than the #NUM! a width check would give.
 */
function toNumeral(
  value: number,
  radix: number,
  bits: number,
  placesArg: Value | undefined,
): string | CellError {
  const span = 2 ** bits;
  if (value < -span / 2 || value >= span / 2) return CellError.NUM;
  const text = (value < 0 ? value + span : value).toString(radix).toUpperCase();

  if (placesArg === undefined || scalarArg(placesArg) === null) return text;
  // The type check happens even when the value is negative: Excel converts the
  // argument before it decides to ignore it, so DEC2BIN(-1,"x") is #VALUE!
  // while DEC2BIN(-1,2) is the full ten digits.
  const places = engInteger(placesArg);
  if (isError(places)) return places;
  if (value < 0) return text;
  // Ten digits is the whole representable width, so a wider request has no
  // meaning and Excel refuses it rather than padding past the sign bit.
  if (places <= 0 || places > 10 || places < text.length) return CellError.NUM;
  return text.padStart(places, '0');
}

/** DEC2BIN and its siblings: a decimal input, encoded to a fixed width. */
function decToBase(radix: number, bits: number): FunctionSpec['impl'] {
  return (args) => {
    const n = engInteger(args[0]);
    if (isError(n)) return n;
    return toNumeral(n, radix, bits, args[1]);
  };
}

/** BIN2DEC and its siblings: a fixed-width input, decoded to a number. */
function baseToDec(radix: number, bits: number): FunctionSpec['impl'] {
  return (args) => {
    const text = numeralArg(args[0]);
    if (isError(text)) return text;
    return fromNumeral(text, radix, bits);
  };
}

/** BIN2HEX and the other cross conversions, which re-encode at a new width. */
function baseToBase(
  fromRadix: number,
  fromBits: number,
  toRadix: number,
  toBits: number,
): FunctionSpec['impl'] {
  return (args) => {
    const text = numeralArg(args[0]);
    if (isError(text)) return text;
    const value = fromNumeral(text, fromRadix, fromBits);
    if (isError(value)) return value;
    return toNumeral(value, toRadix, toBits, args[1]);
  };
}

function conversionSpec(
  name: string,
  first: string,
  impl: FunctionSpec['impl'],
  summary: string,
): FunctionSpec {
  return {
    name,
    params: [p.scalar(first), p.scalar('places', true)],
    broadcast: true,
    summary,
    impl,
  };
}

const BIN2DEC = conversionSpec(
  'BIN2DEC', 'number', baseToDec(2, BIN_BITS),
  'A binary number, as a decimal number.',
);
const BIN2OCT = conversionSpec(
  'BIN2OCT', 'number', baseToBase(2, BIN_BITS, 8, OCT_BITS),
  'A binary number, as an octal number.',
);
const BIN2HEX = conversionSpec(
  'BIN2HEX', 'number', baseToBase(2, BIN_BITS, 16, HEX_BITS),
  'A binary number, as a hexadecimal number.',
);
const OCT2DEC = conversionSpec(
  'OCT2DEC', 'number', baseToDec(8, OCT_BITS),
  'An octal number, as a decimal number.',
);
const OCT2BIN = conversionSpec(
  'OCT2BIN', 'number', baseToBase(8, OCT_BITS, 2, BIN_BITS),
  'An octal number, as a binary number.',
);
const OCT2HEX = conversionSpec(
  'OCT2HEX', 'number', baseToBase(8, OCT_BITS, 16, HEX_BITS),
  'An octal number, as a hexadecimal number.',
);
const HEX2DEC = conversionSpec(
  'HEX2DEC', 'number', baseToDec(16, HEX_BITS),
  'A hexadecimal number, as a decimal number.',
);
const HEX2BIN = conversionSpec(
  'HEX2BIN', 'number', baseToBase(16, HEX_BITS, 2, BIN_BITS),
  'A hexadecimal number, as a binary number.',
);
const HEX2OCT = conversionSpec(
  'HEX2OCT', 'number', baseToBase(16, HEX_BITS, 8, OCT_BITS),
  'A hexadecimal number, as an octal number.',
);
const DEC2BIN = conversionSpec(
  'DEC2BIN', 'number', decToBase(2, BIN_BITS),
  'A decimal number, as a binary number.',
);
const DEC2OCT = conversionSpec(
  'DEC2OCT', 'number', decToBase(8, OCT_BITS),
  'A decimal number, as an octal number.',
);
const DEC2HEX = conversionSpec(
  'DEC2HEX', 'number', decToBase(16, HEX_BITS),
  'A decimal number, as a hexadecimal number.',
);

// ---------------------------------------------------------------------------
// Bitwise
// ---------------------------------------------------------------------------

/** Excel's bitwise functions work on 48-bit unsigned integers. */
const BIT_MAX = 2 ** 48 - 1;
const BIT_SPLIT = 2 ** 24;

function bitOperand(v: Value | undefined): number | CellError {
  const n = engNumber(v);
  if (isError(n)) return n;
  if (n < 0 || n > BIT_MAX || !Number.isInteger(n)) return CellError.NUM;
  return n;
}

/**
 * Apply a 32-bit operation to a 48-bit value by halving it.
 *
 * JavaScript's bitwise operators truncate to 32 bits, which would silently drop
 * the top sixteen bits of a legal BITAND argument.
 */
function bitwise(a: number, b: number, op: (x: number, y: number) => number): number {
  const high = op(Math.floor(a / BIT_SPLIT), Math.floor(b / BIT_SPLIT));
  const low = op(a % BIT_SPLIT, b % BIT_SPLIT);
  return high * BIT_SPLIT + low;
}

function bitSpec(name: string, op: (x: number, y: number) => number, summary: string): FunctionSpec {
  return {
    name,
    params: [p.scalar('number1'), p.scalar('number2')],
    broadcast: true,
    summary,
    impl: (args) => {
      const a = bitOperand(args[0]);
      if (isError(a)) return a;
      const b = bitOperand(args[1]);
      if (isError(b)) return b;
      return bitwise(a, b, op);
    },
  };
}

const BITAND = bitSpec('BITAND', (x, y) => x & y, 'A bitwise AND of two numbers.');
const BITOR = bitSpec('BITOR', (x, y) => x | y, 'A bitwise OR of two numbers.');
const BITXOR = bitSpec('BITXOR', (x, y) => x ^ y, 'A bitwise exclusive OR of two numbers.');

/** A shift by any amount up to 53 bits; a negative amount shifts the other way. */
function shift(value: number, amount: number): number | CellError {
  if (Math.abs(amount) > 53) return CellError.NUM;
  const moved = amount >= 0 ? value * 2 ** amount : Math.floor(value / 2 ** -amount);
  return moved > BIT_MAX ? CellError.NUM : moved;
}

const BITLSHIFT: FunctionSpec = {
  name: 'BITLSHIFT',
  params: [p.scalar('number'), p.scalar('shift_amount')],
  broadcast: true,
  summary: 'A number shifted left by a number of bits.',
  impl: (args) => {
    const value = bitOperand(args[0]);
    if (isError(value)) return value;
    const amount = engNumber(args[1]);
    if (isError(amount)) return amount;
    if (!Number.isInteger(amount)) return CellError.NUM;
    return shift(value, amount);
  },
};

const BITRSHIFT: FunctionSpec = {
  name: 'BITRSHIFT',
  params: [p.scalar('number'), p.scalar('shift_amount')],
  broadcast: true,
  summary: 'A number shifted right by a number of bits.',
  impl: (args) => {
    const value = bitOperand(args[0]);
    if (isError(value)) return value;
    const amount = engNumber(args[1]);
    if (isError(amount)) return amount;
    if (!Number.isInteger(amount)) return CellError.NUM;
    return shift(value, -amount);
  },
};

// ---------------------------------------------------------------------------
// Complex numbers
// ---------------------------------------------------------------------------

/**
 * A parsed complex number.
 *
 * The suffix is empty when the text carried no imaginary part at all, which is
 * how a real argument stays compatible with both an i-suffixed and a
 * j-suffixed operand.
 */
interface Complex {
  re: number;
  im: number;
  suffix: '' | 'i' | 'j';
}

/** The strict numeric grammar the IM family accepts: no percent, no grouping. */
const COMPLEX_NUMBER = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

function parseComplexNumber(text: string): number | undefined {
  if (!COMPLEX_NUMBER.test(text)) return undefined;
  const n = Number(text);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Parse Excel's x+yi text form.
 *
 * The split point is the last sign that is neither the leading sign nor part of
 * an exponent, so 1.5e-3+2e+4i separates where a reader would put it.
 */
function parseComplexText(text: string): Complex | CellError {
  const t = text.trim();
  if (t.length === 0) return CellError.NUM;

  const last = t[t.length - 1]!;
  if (last !== 'i' && last !== 'j') {
    const re = parseComplexNumber(t);
    return re === undefined ? CellError.NUM : { re, im: 0, suffix: '' };
  }

  const body = t.slice(0, -1);
  let split = -1;
  for (let i = body.length - 1; i > 0; i--) {
    const ch = body[i]!;
    if (ch !== '+' && ch !== '-') continue;
    const before = body[i - 1]!;
    if (before === 'e' || before === 'E') continue;
    split = i;
    break;
  }

  const realText = split < 0 ? '' : body.slice(0, split);
  const imagText = split < 0 ? body : body.slice(split);

  const re = realText === '' ? 0 : parseComplexNumber(realText);
  if (re === undefined) return CellError.NUM;
  // A bare suffix means one: 3+i is 3+1i, and -i is -1i.
  const im =
    imagText === '' || imagText === '+' ? 1 : imagText === '-' ? -1 : parseComplexNumber(imagText);
  if (im === undefined) return CellError.NUM;
  return { re, im, suffix: last };
}

function toComplex(v: Value | undefined): Complex | CellError {
  const s = scalarArg(v);
  if (isError(s)) return s;
  if (typeof s === 'boolean') return CellError.VALUE;
  if (s === null) return { re: 0, im: 0, suffix: '' };
  if (typeof s === 'number') return { re: s, im: 0, suffix: '' };
  return parseComplexText(s);
}

/** The suffix a result should carry, or #VALUE! when the operands disagree. */
function mergeSuffix(parts: readonly Complex[]): '' | 'i' | 'j' | CellError {
  let found: '' | 'i' | 'j' = '';
  for (const part of parts) {
    if (part.suffix === '') continue;
    if (found !== '' && found !== part.suffix) return CellError.VALUE;
    found = part.suffix;
  }
  return found;
}

/** Excel prints an integral part without a decimal point and -0 as 0. */
function formatPart(v: number): string {
  return formatNumberForConcat(v === 0 ? 0 : v);
}

function formatComplex(re: number, im: number, suffix: '' | 'i' | 'j'): string | CellError {
  if (!Number.isFinite(re) || !Number.isFinite(im)) return CellError.NUM;
  const unit = suffix === '' ? 'i' : suffix;
  if (im === 0) return formatPart(re);

  const magnitude = Math.abs(im) === 1 ? '' : formatPart(Math.abs(im));
  const sign = im < 0 ? '-' : '+';
  if (re === 0) return `${im < 0 ? '-' : ''}${magnitude}${unit}`;
  return `${formatPart(re)}${sign}${magnitude}${unit}`;
}

/** A unary IM function: parse, transform, re-format with the same suffix. */
function complexSpec(
  name: string,
  summary: string,
  fn: (z: Complex) => { re: number; im: number } | CellError,
): FunctionSpec {
  return {
    name,
    params: [p.scalar('inumber')],
    broadcast: true,
    summary,
    impl: (args) => {
      const z = toComplex(args[0]);
      if (isError(z)) return z;
      const out = fn(z);
      if (isError(out)) return out;
      return formatComplex(out.re, out.im, z.suffix);
    },
  };
}

/** An IM function returning a real number rather than a complex one. */
function complexScalarSpec(
  name: string,
  summary: string,
  fn: (z: Complex) => number | CellError,
): FunctionSpec {
  return {
    name,
    params: [p.scalar('inumber')],
    broadcast: true,
    summary,
    impl: (args) => {
      const z = toComplex(args[0]);
      if (isError(z)) return z;
      const out = fn(z);
      return isError(out) ? out : checkMagnitude(out);
    },
  };
}

function mul(a: Complex, b: Complex): { re: number; im: number } {
  return { re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re };
}

function div(a: Complex, b: Complex): { re: number; im: number } | CellError {
  const denom = b.re * b.re + b.im * b.im;
  if (denom === 0) return CellError.NUM;
  return {
    re: (a.re * b.re + a.im * b.im) / denom,
    im: (a.im * b.re - a.re * b.im) / denom,
  };
}

/** The principal logarithm, which IMLN, IMLOG2, IMLOG10 and IMPOWER all need. */
function complexLog(z: Complex): { re: number; im: number } | CellError {
  const r = Math.hypot(z.re, z.im);
  if (r === 0) return CellError.NUM;
  return { re: Math.log(r), im: Math.atan2(z.im, z.re) };
}

const COMPLEX: FunctionSpec = {
  name: 'COMPLEX',
  params: [p.scalar('real_num'), p.scalar('i_num'), p.scalar('suffix', true)],
  broadcast: true,
  summary: 'A complex number from its real and imaginary coefficients.',
  impl: (args) => {
    const re = engNumber(args[0]);
    if (isError(re)) return re;
    const im = engNumber(args[1]);
    if (isError(im)) return im;
    const raw = scalarArg(args[2]);
    if (isError(raw)) return raw;
    // Excel is case-sensitive here: an upper-case I is not a valid suffix.
    const suffix = raw === null ? 'i' : raw;
    if (suffix !== 'i' && suffix !== 'j') return CellError.VALUE;
    return formatComplex(re, im, suffix);
  },
};

const IMREAL = complexScalarSpec('IMREAL', 'The real coefficient of a complex number.', (z) => z.re);
const IMAGINARY = complexScalarSpec(
  'IMAGINARY', 'The imaginary coefficient of a complex number.', (z) => z.im,
);
const IMABS = complexScalarSpec(
  'IMABS', 'The absolute value of a complex number.', (z) => Math.hypot(z.re, z.im),
);
const IMARGUMENT = complexScalarSpec(
  'IMARGUMENT',
  'The argument theta of a complex number, in radians.',
  // Zero has no argument, and Excel reports that as a division by zero rather
  // than as #NUM!.
  (z) => (z.re === 0 && z.im === 0 ? CellError.DIV0 : Math.atan2(z.im, z.re)),
);
const IMCONJUGATE = complexSpec(
  'IMCONJUGATE', 'The complex conjugate of a complex number.', (z) => ({ re: z.re, im: -z.im }),
);

/** IMSUM and IMPRODUCT, which take any number of operands including ranges. */
function complexFoldSpec(
  name: string,
  summary: string,
  seed: { re: number; im: number },
  step: (acc: Complex, next: Complex) => { re: number; im: number },
): FunctionSpec {
  return {
    name,
    // The operands are taken as arrays rather than scalars so that IMSUM over a
    // range sums the range instead of its top-left cell.
    params: [p.array('inumber1'), p.rest('inumber')],
    summary,
    impl: (args) => {
      const parts: Complex[] = [];
      for (const cell of flatten(args)) {
        const z = toComplex(cell);
        if (isError(z)) return z;
        parts.push(z);
      }
      const suffix = mergeSuffix(parts);
      if (isError(suffix)) return suffix;
      let acc = { ...seed, suffix } as Complex;
      for (const part of parts) acc = { ...step(acc, part), suffix };
      return formatComplex(acc.re, acc.im, suffix);
    },
  };
}

const IMSUM = complexFoldSpec(
  'IMSUM', 'The sum of complex numbers.', { re: 0, im: 0 },
  (acc, next) => ({ re: excelAdd(acc.re, next.re), im: excelAdd(acc.im, next.im) }),
);
const IMPRODUCT = complexFoldSpec(
  'IMPRODUCT', 'The product of complex numbers.', { re: 1, im: 0 }, mul,
);

/** A binary IM function, which must agree on the suffix before it computes. */
function complexBinarySpec(
  name: string,
  summary: string,
  fn: (a: Complex, b: Complex) => { re: number; im: number } | CellError,
): FunctionSpec {
  return {
    name,
    params: [p.scalar('inumber1'), p.scalar('inumber2')],
    broadcast: true,
    summary,
    impl: (args) => {
      const a = toComplex(args[0]);
      if (isError(a)) return a;
      const b = toComplex(args[1]);
      if (isError(b)) return b;
      const suffix = mergeSuffix([a, b]);
      if (isError(suffix)) return suffix;
      const out = fn(a, b);
      if (isError(out)) return out;
      return formatComplex(out.re, out.im, suffix);
    },
  };
}

const IMSUB = complexBinarySpec(
  'IMSUB', 'The difference of two complex numbers.',
  (a, b) => ({ re: a.re - b.re, im: a.im - b.im }),
);
const IMDIV = complexBinarySpec('IMDIV', 'The quotient of two complex numbers.', div);

const IMPOWER: FunctionSpec = {
  name: 'IMPOWER',
  params: [p.scalar('inumber'), p.scalar('number')],
  broadcast: true,
  summary: 'A complex number raised to a power.',
  impl: (args) => {
    const z = toComplex(args[0]);
    if (isError(z)) return z;
    const n = engNumber(args[1]);
    if (isError(n)) return n;
    const r = Math.hypot(z.re, z.im);
    if (r === 0) {
      if (n === 0) return formatComplex(1, 0, z.suffix);
      return n < 0 ? CellError.NUM : formatComplex(0, 0, z.suffix);
    }
    const magnitude = r ** n;
    const angle = Math.atan2(z.im, z.re) * n;
    return formatComplex(magnitude * Math.cos(angle), magnitude * Math.sin(angle), z.suffix);
  },
};

const IMSQRT = complexSpec('IMSQRT', 'The square root of a complex number.', (z) => {
  const r = Math.sqrt(Math.hypot(z.re, z.im));
  const angle = Math.atan2(z.im, z.re) / 2;
  return { re: r * Math.cos(angle), im: r * Math.sin(angle) };
});

const IMEXP = complexSpec('IMEXP', 'The exponential of a complex number.', (z) => {
  const scale = Math.exp(z.re);
  return { re: scale * Math.cos(z.im), im: scale * Math.sin(z.im) };
});

const IMLN = complexSpec('IMLN', 'The natural logarithm of a complex number.', complexLog);

function scaledLog(divisor: number): (z: Complex) => { re: number; im: number } | CellError {
  return (z) => {
    const ln = complexLog(z);
    if (isError(ln)) return ln;
    return { re: ln.re / divisor, im: ln.im / divisor };
  };
}

const IMLOG2 = complexSpec(
  'IMLOG2', 'The base-2 logarithm of a complex number.', scaledLog(Math.LN2),
);
const IMLOG10 = complexSpec(
  'IMLOG10', 'The base-10 logarithm of a complex number.', scaledLog(Math.LN10),
);

function sine(z: Complex): { re: number; im: number } {
  return { re: Math.sin(z.re) * Math.cosh(z.im), im: Math.cos(z.re) * Math.sinh(z.im) };
}

function cosine(z: Complex): { re: number; im: number } {
  return { re: Math.cos(z.re) * Math.cosh(z.im), im: -Math.sin(z.re) * Math.sinh(z.im) };
}

const IMSIN = complexSpec('IMSIN', 'The sine of a complex number.', sine);
const IMCOS = complexSpec('IMCOS', 'The cosine of a complex number.', cosine);

const IMSINH = complexSpec('IMSINH', 'The hyperbolic sine of a complex number.', (z) => ({
  re: Math.sinh(z.re) * Math.cos(z.im),
  im: Math.cosh(z.re) * Math.sin(z.im),
}));

const IMCOSH = complexSpec('IMCOSH', 'The hyperbolic cosine of a complex number.', (z) => ({
  re: Math.cosh(z.re) * Math.cos(z.im),
  im: Math.sinh(z.re) * Math.sin(z.im),
}));

/**
 * The four remaining circular functions are written over the double angle -
 * tan(a+bi) = (sin 2a + i sinh 2b)/(cos 2a + cosh 2b), and its relatives - not
 * as divisions of IMSIN by IMCOS. The two agree to within a unit in the last
 * place, and this is the form whose last place agrees with Excel's.
 */
const IMTAN = complexSpec('IMTAN', 'The tangent of a complex number.', (z) => {
  const denominator = Math.cos(2 * z.re) + Math.cosh(2 * z.im);
  return { re: Math.sin(2 * z.re) / denominator, im: Math.sinh(2 * z.im) / denominator };
});

// IMCOT arrived in Excel 2013 alongside IMSEC and IMCSC, so it is stored
// prefixed; the registry's FUTURE_FUNCTIONS list now says so, which is what the
// xlsx writer consults.
const IMCOT = complexSpec('IMCOT', 'The cotangent of a complex number.', (z) => {
  const denominator = Math.cosh(2 * z.im) - Math.cos(2 * z.re);
  return { re: Math.sin(2 * z.re) / denominator, im: -Math.sinh(2 * z.im) / denominator };
});

const IMSEC = complexSpec('IMSEC', 'The secant of a complex number.', (z) => {
  const denominator = Math.cos(2 * z.re) + Math.cosh(2 * z.im);
  return {
    re: (2 * Math.cos(z.re) * Math.cosh(z.im)) / denominator,
    im: (2 * Math.sin(z.re) * Math.sinh(z.im)) / denominator,
  };
});

const IMCSC = complexSpec('IMCSC', 'The cosecant of a complex number.', (z) => {
  const denominator = Math.cosh(2 * z.im) - Math.cos(2 * z.re);
  return {
    re: (2 * Math.sin(z.re) * Math.cosh(z.im)) / denominator,
    im: (-2 * Math.cos(z.re) * Math.sinh(z.im)) / denominator,
  };
});

// ---------------------------------------------------------------------------
// The error function
// ---------------------------------------------------------------------------

const TWO_OVER_SQRT_PI = 2 / Math.sqrt(Math.PI);
const ONE_OVER_SQRT_PI = 1 / Math.sqrt(Math.PI);

/**
 * erf for small arguments, from the exponentially scaled ascending series.
 *
 * The plain Maclaurin series alternates and loses most of its digits by x = 3;
 * this form, erf(x) = (2x/sqrt(pi)) e^-x^2 sum (2x^2)^n / (1.3.5...(2n+1)), has
 * only positive terms and so keeps them all.
 */
function erfSeries(x: number): number {
  const xx = x * x;
  let term = 1;
  let sum = 1;
  for (let n = 1; n < 200; n++) {
    term *= (2 * xx) / (2 * n + 1);
    sum += term;
    if (term < sum * 1e-17) break;
  }
  return TWO_OVER_SQRT_PI * x * Math.exp(-xx) * sum;
}

/**
 * erfc for large arguments, from the continued fraction
 * erfc(x) = e^-x^2/sqrt(pi) / (x + (1/2)/(x + 1/(x + (3/2)/(x + ...)))),
 * evaluated by the modified Lentz algorithm.
 */
function erfcFraction(x: number): number {
  const tiny = 1e-300;
  let f = tiny;
  let c = f;
  let d = 0;
  for (let n = 0; n < 300; n++) {
    const a = n === 0 ? 1 : n / 2;
    d = x + a * d;
    if (d === 0) d = tiny;
    c = x + a / c;
    if (c === 0) c = tiny;
    d = 1 / d;
    const delta = c * d;
    f *= delta;
    // One unit in the last place: the convergents of this fraction close in on
    // the limit steadily, so the first step that changes nothing is the end.
    if (Math.abs(delta - 1) < 1e-16) break;
  }
  return ONE_OVER_SQRT_PI * Math.exp(-x * x) * f;
}

/**
 * The crossover: below it the series is short, above it the fraction converges
 * in a few dozen steps. Just below it, erfc is 1 - erf and loses about two
 * digits to the subtraction, which is the worst this pair does anywhere.
 */
const ERF_CROSSOVER = 2;

function erf(x: number): number {
  const ax = Math.abs(x);
  if (ax < ERF_CROSSOVER) return erfSeries(x);
  const tail = 1 - erfcFraction(ax);
  return x < 0 ? -tail : tail;
}

function erfc(x: number): number {
  if (x < 0) return 2 - erfc(-x);
  return x < ERF_CROSSOVER ? 1 - erfSeries(x) : erfcFraction(x);
}

const ERF: FunctionSpec = {
  name: 'ERF',
  params: [p.scalar('lower_limit'), p.scalar('upper_limit', true)],
  broadcast: true,
  summary: 'The error function integrated between two limits.',
  impl: (args) => {
    const lower = engNumber(args[0]);
    if (isError(lower)) return lower;
    if (args[1] === undefined || scalarArg(args[1]) === null) return erf(lower);
    const upper = engNumber(args[1]);
    if (isError(upper)) return upper;
    return erf(upper) - erf(lower);
  },
};

const ERF_PRECISE: FunctionSpec = {
  name: 'ERF.PRECISE',
  params: [p.scalar('x')],
  broadcast: true,
  summary: 'The error function integrated between zero and x.',
  impl: (args) => {
    const x = engNumber(args[0]);
    return isError(x) ? x : erf(x);
  },
};

const ERFC: FunctionSpec = {
  name: 'ERFC',
  params: [p.scalar('x')],
  broadcast: true,
  summary: 'The complementary error function integrated between x and infinity.',
  impl: (args) => {
    const x = engNumber(args[0]);
    return isError(x) ? x : erfc(x);
  },
};

const ERFC_PRECISE: FunctionSpec = {
  ...ERFC,
  name: 'ERFC.PRECISE',
};

// ---------------------------------------------------------------------------
// Bessel functions
// ---------------------------------------------------------------------------

/**
 * Sixteen-point Gauss-Legendre on [-1, 1], as eight symmetric pairs.
 *
 * Sixteen points integrate a polynomial of degree 31 exactly, which on a panel
 * narrow enough to hold about one oscillation of the integrand leaves an error
 * at the rounding level.
 */
const GL_NODES = [
  0.0950125098376374, 0.2816035507792589, 0.4580167776572274, 0.6178762444026438,
  0.7554044083550030, 0.8656312023878318, 0.9445750230732326, 0.9894009349916499,
] as const;
const GL_WEIGHTS = [
  0.1894506104550685, 0.1826034150449236, 0.1691565193950025, 0.1495959888165767,
  0.1246289712555339, 0.0951585116824928, 0.0622535239386479, 0.0271524594117541,
] as const;

/** Composite Gauss-Legendre over [a, b]. */
function integrate(f: (t: number) => number, a: number, b: number, panels: number): number {
  const h = (b - a) / panels;
  const half = h / 2;
  let total = 0;
  for (let k = 0; k < panels; k++) {
    const mid = a + h * (k + 0.5);
    let sum = 0;
    for (let j = 0; j < 8; j++) {
      const offset = half * GL_NODES[j]!;
      sum += GL_WEIGHTS[j]! * (f(mid + offset) + f(mid - offset));
    }
    total += sum * half;
  }
  return total;
}

/**
 * The largest number of kernel evaluations any one Bessel call may spend.
 *
 * Both quadratures below need work proportional to x, so without a ceiling a
 * single cell holding BESSELY(1E9,0) would run for the better part of an hour.
 * Everything past the ceiling is served by the asymptotic expansion instead,
 * which costs nothing and is more accurate there anyway.
 */
const BESSEL_WORK_LIMIT = 1 << 21;

/** The smallest power of two at least `v`, clamped to the work limit. */
function pow2AtLeast(v: number): number {
  let n = 64;
  while (n < v && n < BESSEL_WORK_LIMIT) n *= 2;
  return n;
}

/**
 * The trapezoidal rule over one period of an analytic periodic integrand, which
 * converges geometrically: for the Bessel integrals the truncation error is
 * J_(N-n)(x) + J_(N+n)(x), and once N is comfortably past x + n that is below
 * the rounding level. The node count starts where that is already true and then
 * doubles until two successive answers agree, so the accuracy is checked rather
 * than assumed.
 *
 * The starting count matters: a fixed one silently returns the unconverged
 * answer for a large argument, and J_0(20000) computed on 8192 nodes is not
 * merely inaccurate, it is three times larger than J_0 can ever be.
 */
function periodicMean(kernel: (theta: number) => number, start: number): number {
  let previous = Number.NaN;
  let mean = Number.NaN;
  for (let n = pow2AtLeast(start); n <= BESSEL_WORK_LIMIT; n *= 2) {
    let sum = 0;
    for (let k = 0; k < n; k++) sum += kernel((2 * Math.PI * k) / n);
    mean = sum / n;
    if (Number.isFinite(previous) && Math.abs(mean - previous) <= 1e-16 * Math.max(1, Math.abs(mean))) {
      return mean;
    }
    previous = mean;
  }
  return mean;
}

/**
 * Hankel's asymptotic expansion, which gives J_n and Y_n together:
 *
 *   J_n(x) = sqrt(2/(pi x)) (P cos w - Q sin w),
 *   Y_n(x) = sqrt(2/(pi x)) (P sin w + Q cos w),   w = x - (n/2 + 1/4) pi,
 *
 * with P the even and Q the odd part of the series whose k-th term is
 * prod_(j<=k) (4n^2 - (2j-1)^2) / (k! (8x)^k), signed (-1)^floor(k/2). The
 * series diverges eventually, so it is summed only while its terms shrink; the
 * error is then bounded by the first term dropped, which for x >= 1000 and
 * 4n^2 <= x is far below the rounding level. The only real error left is the
 * phase: w loses the last bits of x, so a result near a zero of J or Y carries
 * an absolute error of about ulp(x) times the amplitude. Excel's own
 * approximations have the same limit and a far worse one besides.
 */
function hankel(x: number, n: number): { j: number; y: number } {
  const mu = 4 * n * n;
  let p = 1;
  let q = 0;
  let term = 1;
  let smallest = Number.POSITIVE_INFINITY;
  for (let k = 1; k <= 200; k++) {
    term *= (mu - (2 * k - 1) ** 2) / (k * 8 * x);
    const size = Math.abs(term);
    if (size > smallest) break;
    smallest = size;
    const sign = k % 4 === 1 || k % 4 === 0 ? 1 : -1;
    if (k % 2 === 1) q += sign * term;
    else p += sign * term;
    if (size < 1e-19) break;
  }
  const phase = x - (n / 2 + 0.25) * Math.PI;
  const amplitude = Math.sqrt(2 / (Math.PI * x));
  const cos = Math.cos(phase);
  const sin = Math.sin(phase);
  return { j: amplitude * (p * cos - q * sin), y: amplitude * (p * sin + q * cos) };
}

/**
 * Where the asymptotic expansion is the better of the two.
 *
 * Its terms only shrink while 4n^2 stays under 8x, and by x = 1000 it agrees
 * with the quadrature to a few parts in 10^14; below that the quadrature is
 * both cheap and exact, so there is nothing to gain by switching earlier.
 */
function asymptoticApplies(x: number, n: number): boolean {
  return x >= 1000 && 4 * n * n <= x;
}

/**
 * The ascending series for J_n, used where the integral would cancel.
 *
 * For x below about 2*sqrt(n+1) the terms decrease from the first one, so the
 * alternating signs cost nothing; that is exactly the regime where J_n(x) is
 * astronomically smaller than the order-one integrand and the quadrature would
 * return rounding noise.
 */
function besselJSeries(x: number, n: number): number {
  const half = x / 2;
  let term = half ** n;
  for (let k = 2; k <= n; k++) term /= k;
  let sum = term;
  for (let k = 1; k < 400; k++) {
    term *= (-half * half) / (k * (n + k));
    sum += term;
    if (Math.abs(term) < Math.abs(sum) * 1e-18) break;
  }
  return sum;
}

function besselJ(x: number, n: number): number {
  // J_n(-x) = (-1)^n J_n(x), so only the positive half needs a method.
  const ax = Math.abs(x);
  const sign = x < 0 && n % 2 === 1 ? -1 : 1;
  if (ax * ax <= 4 * (n + 1)) return sign * besselJSeries(ax, n);
  // The trapezoid needs a node for every radian of phase and then some; past
  // the work ceiling the asymptotic expansion takes over, exactly as it does
  // for Y, rather than the quadrature quietly returning an unconverged answer.
  const nodes = 4 * (ax + n) + 64;
  if (asymptoticApplies(ax, n) || nodes > BESSEL_WORK_LIMIT) return sign * hankel(ax, n).j;
  return sign * periodicMean((theta) => Math.cos(n * theta - ax * Math.sin(theta)), nodes);
}

/**
 * The ascending series for I_n. Every term is positive, so this is
 * unconditionally well conditioned; only the first term needs care, since
 * (x/2)^n / n! overflows long before I_n(x) does.
 */
function besselI(x: number, n: number): number {
  const ax = Math.abs(x);
  const half = ax / 2;
  let logTerm = n * Math.log(half === 0 ? Number.MIN_VALUE : half);
  for (let k = 2; k <= n; k++) logTerm -= Math.log(k);
  let term = half === 0 ? (n === 0 ? 1 : 0) : Math.exp(logTerm);
  let sum = term;
  for (let k = 1; k < 4000; k++) {
    term *= (half * half) / (k * (n + k));
    sum += term;
    if (term < sum * 1e-18) break;
  }
  // I_n(-x) = (-1)^n I_n(x); the series above is written in |x|.
  return x < 0 && n % 2 === 1 ? -sum : sum;
}

/** How many nats below its peak an exponent may fall before it stops mattering. */
const DECAY_NATS = 50;

/**
 * Where to cut off an integrand of the form e^(n*t - x*g(t)).
 *
 * The exponent peaks where x*g'(t) = n and then falls; integrating past the
 * point where it has fallen fifty nats below the peak adds nothing a double can
 * hold. The cut-off has to be tight, not merely sufficient: for x = 20000 the
 * whole of e^(-x sinh t) lives inside t < 0.0025, and a limit rounded up to 0.5
 * hands the quadrature a range four hundred times wider than the integrand,
 * which then integrates the spike with a single node and loses six digits of
 * Y_0. So the crossing is bracketed by doubling and then bisected, which is
 * bounded work and lands on the exact point.
 */
function decayLimit(exponent: (t: number) => number, peak: number): number {
  const top = exponent(peak);
  const target = top - DECAY_NATS;
  const ceiling = 750;
  let offset = 1e-12;
  while (peak + offset < ceiling && exponent(peak + offset) > target) offset *= 2;
  let lo = peak;
  let hi = Math.min(peak + offset, ceiling);
  for (let i = 0; i < 80; i++) {
    const mid = (lo + hi) / 2;
    if (mid <= lo || mid >= hi) break;
    if (exponent(mid) > target) lo = mid;
    else hi = mid;
  }
  return hi;
}

/**
 * Panels enough to resolve an exponential integrand over [0, limit].
 *
 * What has to be resolved is the exponent's swing, not the width of the
 * interval: two nats per panel leaves sixteen Gauss-Legendre nodes with an
 * almost straight integrand to fit.
 */
function decayPanels(exponent: (t: number) => number, peak: number): number {
  const swing = exponent(peak) - exponent(0) + DECAY_NATS;
  return Math.min(4096, Math.max(16, Math.ceil(swing / 2)));
}

/**
 * Y_n from Schlaefli's integral,
 * Y_n(x) = (1/pi) int_0^pi sin(x sin t - n t) dt
 *        - (1/pi) int_0^inf (e^(n t) + (-1)^n e^(-n t)) e^(-x sinh t) dt.
 *
 * The first integrand oscillates through about 2x + n*pi radians of phase, so
 * the panel count grows with both; the second decays double-exponentially and
 * needs only a short range.
 */
function besselY(x: number, n: number): number {
  const panels = Math.max(16, Math.ceil(x + 1.6 * n) + 8);
  // Past the ceiling the oscillatory panel count is the whole cost, and the
  // asymptotic expansion is both free and better.
  if (asymptoticApplies(x, n) || panels * 16 > BESSEL_WORK_LIMIT) return hankel(x, n).y;

  const oscillatory = integrate((t) => Math.sin(x * Math.sin(t) - n * t), 0, Math.PI, panels);

  const sign = n % 2 === 0 ? 1 : -1;
  const exponent = (t: number): number => n * t - x * Math.sinh(t);
  const peak = n > x ? Math.acosh(n / x) : 0;
  const limit = decayLimit(exponent, peak);
  const decaying = integrate(
    (t) => (Math.exp(n * t) + sign * Math.exp(-n * t)) * Math.exp(-x * Math.sinh(t)),
    0,
    limit,
    decayPanels(exponent, peak),
  );

  return (oscillatory - decaying) / Math.PI;
}

/** K_n from its integral, K_n(x) = int_0^inf e^(-x cosh t) cosh(n t) dt. */
function besselK(x: number, n: number): number {
  const exponent = (t: number): number => n * t - x * Math.cosh(t);
  const peak = Math.asinh(n / x);
  const limit = decayLimit(exponent, peak);
  return integrate(
    (t) => Math.exp(-x * Math.cosh(t)) * Math.cosh(n * t),
    0,
    limit,
    decayPanels(exponent, peak),
  );
}

/** Order beyond which the quadratures would cost more than any user wants. */
const MAX_BESSEL_ORDER = 32768;

function besselSpec(
  name: string,
  summary: string,
  positiveOnly: boolean,
  fn: (x: number, n: number) => number,
): FunctionSpec {
  return {
    name,
    params: [p.scalar('x'), p.scalar('n')],
    broadcast: true,
    summary,
    impl: (args) => {
      const x = engNumber(args[0]);
      if (isError(x)) return x;
      const n = engInteger(args[1]);
      if (isError(n)) return n;
      if (n < 0 || n > MAX_BESSEL_ORDER) return CellError.NUM;
      // Y and K have a branch point at the origin and are undefined to its left.
      if (positiveOnly && x <= 0) return CellError.NUM;
      return checkMagnitude(fn(x, n));
    },
  };
}

const BESSELJ = besselSpec('BESSELJ', 'The Bessel function Jn(x).', false, besselJ);
const BESSELI = besselSpec('BESSELI', 'The modified Bessel function In(x).', false, besselI);
const BESSELY = besselSpec('BESSELY', 'The Bessel function Yn(x).', true, besselY);
const BESSELK = besselSpec('BESSELK', 'The modified Bessel function Kn(x).', true, besselK);

// ---------------------------------------------------------------------------
// Step functions
// ---------------------------------------------------------------------------

const DELTA: FunctionSpec = {
  name: 'DELTA',
  params: [p.scalar('number1'), p.scalar('number2', true)],
  broadcast: true,
  summary: 'One when two numbers are equal, zero otherwise.',
  impl: (args) => {
    const a = engNumber(args[0]);
    if (isError(a)) return a;
    const b = engNumber(args[1]);
    if (isError(b)) return b;
    return a === b ? 1 : 0;
  },
};

const GESTEP: FunctionSpec = {
  name: 'GESTEP',
  params: [p.scalar('number'), p.scalar('step', true)],
  broadcast: true,
  summary: 'One when a number is at least the step, zero otherwise.',
  impl: (args) => {
    const value = engNumber(args[0]);
    if (isError(value)) return value;
    const step = engNumber(args[1]);
    if (isError(step)) return step;
    return value >= step ? 1 : 0;
  },
};

// ---------------------------------------------------------------------------
// CONVERT
// ---------------------------------------------------------------------------

/**
 * A unit, as an affine map onto its group's base unit:
 * base = value * scale + offset.
 *
 * Only temperature uses the offset, but carrying it for every unit means the
 * conversion itself has no special case, and a Celsius-to-Fahrenheit conversion
 * goes through exactly the same code path as gram-to-pound.
 */
interface Unit {
  group: string;
  scale: number;
  offset: number;
  /** A metric prefix may be attached. */
  prefixed: boolean;
  /** The power a prefix is raised to, so that a prefix on m^2 scales as a square. */
  dimension: number;
  /** The binary prefixes Ki, Mi and friends apply, which only information units take. */
  binary: boolean;
}

const UNITS = new Map<string, Unit>();

function defineUnit(
  group: string,
  scale: number,
  names: readonly string[],
  options: { prefixed?: boolean; dimension?: number; binary?: boolean; offset?: number } = {},
): void {
  const unit: Unit = {
    group,
    scale,
    offset: options.offset ?? 0,
    prefixed: options.prefixed ?? false,
    dimension: options.dimension ?? 1,
    binary: options.binary ?? false,
  };
  for (const name of names) UNITS.set(name, unit);
}

// Mass, in grams.
defineUnit('mass', 1, ['g'], { prefixed: true });
defineUnit('mass', 14593.9029372064, ['sg']);
defineUnit('mass', 453.59237, ['lbm']);
defineUnit('mass', 1.66053886e-24, ['u'], { prefixed: true });
defineUnit('mass', 28.349523125, ['ozm']);
defineUnit('mass', 0.06479891, ['grain']);
defineUnit('mass', 45359.237, ['cwt', 'shweight']);
defineUnit('mass', 50802.34544, ['uk_cwt', 'lcwt', 'hweight']);
defineUnit('mass', 6350.29318, ['stone']);
defineUnit('mass', 907184.74, ['ton']);
defineUnit('mass', 1016046.9088, ['uk_ton', 'LTON', 'brton']);

// Distance, in metres.
defineUnit('distance', 1, ['m'], { prefixed: true });
defineUnit('distance', 1609.344, ['mi']);
defineUnit('distance', 1852, ['Nmi']);
defineUnit('distance', 0.0254, ['in']);
defineUnit('distance', 0.3048, ['ft']);
defineUnit('distance', 0.9144, ['yd']);
defineUnit('distance', 1e-10, ['ang'], { prefixed: true });
defineUnit('distance', 1.143, ['ell']);
defineUnit('distance', 9.4607304725808e15, ['ly']);
defineUnit('distance', 3.08567758128e16, ['parsec', 'pc']);
defineUnit('distance', 0.0254 / 72, ['Picapt', 'Pica']);
defineUnit('distance', 0.0254 / 6, ['pica']);
defineUnit('distance', 1609.3472186944374, ['survey_mi']);

// Time, in seconds.
defineUnit('time', 31557600, ['yr']);
defineUnit('time', 86400, ['day', 'd']);
defineUnit('time', 3600, ['hr']);
defineUnit('time', 60, ['mn', 'min']);
defineUnit('time', 1, ['sec', 's'], { prefixed: true });

// Pressure, in pascals.
defineUnit('pressure', 1, ['Pa', 'p'], { prefixed: true });
defineUnit('pressure', 101325, ['atm', 'at'], { prefixed: true });
defineUnit('pressure', 133.322, ['mmHg'], { prefixed: true });
defineUnit('pressure', 6894.75729316836, ['psi']);
defineUnit('pressure', 101325 / 760, ['Torr']);

// Force, in newtons.
defineUnit('force', 1, ['N'], { prefixed: true });
defineUnit('force', 1e-5, ['dyn', 'dy'], { prefixed: true });
defineUnit('force', 4.4482216152605, ['lbf']);
defineUnit('force', 9.80665e-3, ['pond'], { prefixed: true });

// Energy, in joules.
defineUnit('energy', 1, ['J'], { prefixed: true });
defineUnit('energy', 1e-7, ['e'], { prefixed: true });
defineUnit('energy', 4.184, ['c'], { prefixed: true });
defineUnit('energy', 4.1868, ['cal'], { prefixed: true });
defineUnit('energy', 1.602176487e-19, ['eV', 'ev'], { prefixed: true });
defineUnit('energy', 2684519.53769617, ['HPh', 'hh']);
defineUnit('energy', 3600, ['Wh', 'wh'], { prefixed: true });
defineUnit('energy', 1.3558179483314004, ['flb']);
defineUnit('energy', 1055.05585262, ['BTU', 'btu']);

// Power, in watts.
defineUnit('power', 1, ['W', 'w'], { prefixed: true });
defineUnit('power', 745.69987158227022, ['HP', 'h']);
defineUnit('power', 735.49875, ['PS']);

// Magnetism, in teslas.
defineUnit('magnetism', 1, ['T'], { prefixed: true });
defineUnit('magnetism', 1e-4, ['ga'], { prefixed: true });

// Temperature, in kelvins. These are the units the affine offset exists for.
defineUnit('temperature', 1, ['K', 'kel'], { prefixed: true });
defineUnit('temperature', 1, ['C', 'cel'], { prefixed: true, offset: 273.15 });
defineUnit('temperature', 5 / 9, ['F', 'fah'], { prefixed: true, offset: 273.15 - 32 * (5 / 9) });
defineUnit('temperature', 5 / 9, ['Rank']);
defineUnit('temperature', 1.25, ['Reau'], { offset: 273.15 });

// Volume, in cubic metres.
defineUnit('volume', 1, ['m3', 'm^3'], { prefixed: true, dimension: 3 });
defineUnit('volume', 0.001, ['l', 'L', 'lt'], { prefixed: true });
defineUnit('volume', 4.92892159375e-6, ['tsp']);
defineUnit('volume', 5e-6, ['tspm']);
defineUnit('volume', 1.478676478125e-5, ['tbs']);
defineUnit('volume', 2.95735295625e-5, ['oz']);
defineUnit('volume', 2.365882365e-4, ['cup']);
defineUnit('volume', 4.73176473e-4, ['pt', 'us_pt']);
defineUnit('volume', 5.6826125e-4, ['uk_pt']);
defineUnit('volume', 9.46352946e-4, ['qt']);
defineUnit('volume', 1.1365225e-3, ['uk_qt']);
defineUnit('volume', 3.785411784e-3, ['gal']);
defineUnit('volume', 4.54609e-3, ['uk_gal']);
defineUnit('volume', 0.158987294928, ['barrel']);
defineUnit('volume', 0.03523907016688, ['bushel']);
defineUnit('volume', 0.028316846592, ['ft3', 'ft^3']);
defineUnit('volume', 1.6387064e-5, ['in3', 'in^3']);
defineUnit('volume', 0.764554857984, ['yd3', 'yd^3']);
defineUnit('volume', 1609.344 ** 3, ['mi3', 'mi^3']);
defineUnit('volume', 1852 ** 3, ['Nmi3', 'Nmi^3']);
defineUnit('volume', 1e-30, ['ang3', 'ang^3'], { prefixed: true, dimension: 3 });
defineUnit('volume', 9.4607304725808e15 ** 3, ['ly3', 'ly^3']);
defineUnit('volume', (0.0254 / 72) ** 3, ['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3']);
defineUnit('volume', 2.8316846592, ['GRT', 'regton']);
defineUnit('volume', 1.13267386368, ['MTON']);

// Area, in square metres.
defineUnit('area', 1, ['m2', 'm^2'], { prefixed: true, dimension: 2 });
defineUnit('area', 1e-20, ['ang2', 'ang^2'], { prefixed: true, dimension: 2 });
defineUnit('area', 100, ['ar'], { prefixed: true });
defineUnit('area', 0.09290304, ['ft2', 'ft^2']);
defineUnit('area', 10000, ['ha']);
defineUnit('area', 6.4516e-4, ['in2', 'in^2']);
defineUnit('area', 9.4607304725808e15 ** 2, ['ly2', 'ly^2']);
defineUnit('area', 2589988.110336, ['mi2', 'mi^2']);
defineUnit('area', 1852 ** 2, ['Nmi2', 'Nmi^2']);
defineUnit('area', (0.0254 / 72) ** 2, ['Picapt2', 'Picapt^2', 'Pica2', 'Pica^2']);
defineUnit('area', 2500, ['Morgen']);
defineUnit('area', 4046.8564224, ['uk_acre']);
defineUnit('area', 4046.87261, ['us_acre']);
defineUnit('area', 0.83612736, ['yd2', 'yd^2']);

// Information, in bits.
defineUnit('information', 1, ['bit'], { prefixed: true, binary: true });
defineUnit('information', 8, ['byte'], { prefixed: true, binary: true });

// Speed, in metres per second.
defineUnit('speed', 1, ['m/s', 'm/sec'], { prefixed: true });
defineUnit('speed', 1 / 3600, ['m/h', 'm/hr'], { prefixed: true });
defineUnit('speed', 1852 / 3600, ['kn']);
defineUnit('speed', 1853.184 / 3600, ['admkn']);
defineUnit('speed', 0.44704, ['mph']);

const DECIMAL_PREFIXES: readonly (readonly [string, number])[] = [
  ['Y', 1e24], ['Z', 1e21], ['E', 1e18], ['P', 1e15], ['T', 1e12], ['G', 1e9],
  ['M', 1e6], ['k', 1e3], ['h', 1e2], ['da', 1e1], ['e', 1e1], ['d', 1e-1],
  ['c', 1e-2], ['m', 1e-3], ['u', 1e-6], ['µ', 1e-6], ['n', 1e-9],
  ['p', 1e-12], ['f', 1e-15], ['a', 1e-18], ['z', 1e-21], ['y', 1e-24],
];

const BINARY_PREFIXES: readonly (readonly [string, number])[] = [
  ['Yi', 2 ** 80], ['Zi', 2 ** 70], ['Ei', 2 ** 60], ['Pi', 2 ** 50],
  ['Ti', 2 ** 40], ['Gi', 2 ** 30], ['Mi', 2 ** 20], ['ki', 2 ** 10],
];

/**
 * Resolve a unit name, trying the exact spelling before any prefix.
 *
 * Order matters: `kn` is a knot, not a kilonewton, and only the exact-first rule
 * gets that right without a list of exceptions.
 */
function resolveUnit(name: string): { unit: Unit; multiplier: number } | undefined {
  const exact = UNITS.get(name);
  if (exact) return { unit: exact, multiplier: 1 };

  // Binary prefixes are tried first because Ki and Mi would otherwise be read as
  // kilo and mega followed by a stray i.
  for (const [binary, table] of [[true, BINARY_PREFIXES], [false, DECIMAL_PREFIXES]] as const) {
    for (const [prefix, factor] of table) {
      if (!name.startsWith(prefix)) continue;
      const unit = UNITS.get(name.slice(prefix.length));
      if (!unit || !unit.prefixed) continue;
      if (binary && !unit.binary) continue;
      return { unit, multiplier: factor ** unit.dimension };
    }
  }
  return undefined;
}

const CONVERT: FunctionSpec = {
  name: 'CONVERT',
  params: [p.scalar('number'), p.scalar('from_unit'), p.scalar('to_unit')],
  broadcast: true,
  summary: 'A number converted from one measurement system to another.',
  impl: (args) => {
    const value = engNumber(args[0]);
    if (isError(value)) return value;
    const fromName = toText(scalarArg(args[1]));
    if (isError(fromName)) return fromName;
    const toName = toText(scalarArg(args[2]));
    if (isError(toName)) return toName;

    const from = resolveUnit(fromName);
    const to = resolveUnit(toName);
    // An unknown unit and a pair from different groups are the same failure to
    // Excel: #N/A rather than #VALUE!.
    if (!from || !to || from.unit.group !== to.unit.group) return CellError.NA;

    const base = value * from.multiplier * from.unit.scale + from.unit.offset;
    return checkMagnitude((base - to.unit.offset) / to.unit.scale / to.multiplier);
  },
};

// ---------------------------------------------------------------------------
// Database functions
// ---------------------------------------------------------------------------

function asArray(v: Value | undefined, ctx: FunctionContext): ArrayValue | CellError {
  if (v === undefined || v === null) return CellError.VALUE;
  if (isArray(v)) return v;
  if (isRef(v)) return ctx.deref(v);
  if (isError(v)) return v;
  return makeArray(1, 1, [v]);
}

/** Header text, normalised the way Excel matches field names: trimmed, caseless. */
function headerKey(v: Scalar): string | undefined {
  if (v === null) return undefined;
  const text = toText(v);
  if (isError(text)) return undefined;
  const trimmed = text.trim();
  return trimmed === '' ? undefined : trimmed.toUpperCase();
}

/** The zero-based column a `field` argument names, or the error Excel reports. */
function resolveField(field: Scalar, database: ArrayValue): number | CellError {
  if (isError(field)) return field;
  if (typeof field === 'number' || typeof field === 'boolean') {
    const index = Math.trunc(typeof field === 'number' ? field : field ? 1 : 0);
    if (index < 1 || index > database.cols) return CellError.VALUE;
    return index - 1;
  }
  const wanted = headerKey(field);
  if (wanted === undefined) return CellError.VALUE;
  for (let col = 0; col < database.cols; col++) {
    if (headerKey(arrayAt(database, 0, col)) === wanted) return col;
  }
  return CellError.VALUE;
}

/**
 * One criteria cell, compiled.
 *
 * A bare text criterion is a prefix match in a criteria range - "Sci" selects
 * Science - which is where the D functions part company with COUNTIF's
 * whole-value equality. Comparisons, numbers and wildcard patterns go through
 * the shared criterion machinery unchanged.
 */
interface FieldTest {
  column: number;
  criterion: Criterion;
  prefix?: string;
}

function compileCriteria(
  criteria: ArrayValue,
  database: ArrayValue,
): FieldTest[][] | CellError {
  const columns: (number | undefined)[] = [];
  for (let col = 0; col < criteria.cols; col++) {
    const header = headerKey(arrayAt(criteria, 0, col));
    if (header === undefined) {
      // A blank header marks a column that is not a constraint at all, which is
      // how a criteria range wider than its used columns still works.
      columns.push(undefined);
      continue;
    }
    let found: number | undefined;
    for (let dbCol = 0; dbCol < database.cols; dbCol++) {
      if (headerKey(arrayAt(database, 0, dbCol)) === header) {
        found = dbCol;
        break;
      }
    }
    if (found === undefined) return CellError.VALUE;
    columns.push(found);
  }

  const rows: FieldTest[][] = [];
  for (let row = 1; row < criteria.rows; row++) {
    const tests: FieldTest[] = [];
    for (let col = 0; col < criteria.cols; col++) {
      const column = columns[col];
      if (column === undefined) continue;
      const raw = arrayAt(criteria, row, col);
      if (raw === null || raw === '') continue;
      if (isError(raw)) return raw;
      const criterion = parseCriterion(raw);
      const test: FieldTest = { column, criterion };
      if (criterion.op === '=' && typeof criterion.value === 'string' && !criterion.pattern) {
        test.prefix = criterion.value.toUpperCase();
      }
      tests.push(test);
    }
    rows.push(tests);
  }
  return rows;
}

function matchesTest(value: Scalar, test: FieldTest): boolean {
  if (test.prefix !== undefined) {
    const text = toText(value);
    if (isError(text)) return false;
    return text.toUpperCase().startsWith(test.prefix);
  }
  return matchesCriterion(value, test.criterion);
}

/**
 * The rows of the database that satisfy the criteria: rows of the criteria
 * range are alternatives, cells within one row are all required, and a criteria
 * range with no rows below its header selects everything.
 */
function selectRows(database: ArrayValue, criteria: FieldTest[][]): number[] {
  const selected: number[] = [];
  for (let row = 1; row < database.rows; row++) {
    const hit =
      criteria.length === 0 ||
      criteria.some((tests) => tests.every((test) => matchesTest(arrayAt(database, row, test.column), test)));
    if (hit) selected.push(row);
  }
  return selected;
}

/** The values of one field across the matching records. */
function fieldValues(
  args: readonly (Value | undefined)[],
  fieldRequired: boolean,
  ctx: FunctionContext,
): { values: Scalar[]; count: number; omitted: boolean } | CellError {
  const database = asArray(args[0], ctx);
  if (isError(database)) return database;
  // DCOUNT and DCOUNTA may be called with the field left out entirely rather
  // than passed as an empty argument, in which case the criteria arrived
  // second. A range in that position can only be the criteria, since a field is
  // named by a single value.
  const shifted =
    !fieldRequired && args[2] === undefined && (isRef(args[1]) || isArray(args[1]));
  const criteria = asArray(shifted ? args[1] : args[2], ctx);
  if (isError(criteria)) return criteria;
  if (database.rows < 1 || criteria.rows < 1) return CellError.VALUE;

  const compiled = compileCriteria(criteria, database);
  if (isError(compiled)) return compiled;
  const rows = selectRows(database, compiled);

  // The field is declared ArgKind.Any so that a range reaching this position is
  // still recognisable as one; a reference to a cell holding the field name is
  // read here rather than by the evaluator.
  const raw = args[1];
  const fieldArg = shifted ? null : isRef(raw) ? ctx.getScalar(raw.sheet, raw.startRow, raw.startCol) : scalarArg(raw);
  if (fieldArg === null) {
    if (fieldRequired) return CellError.VALUE;
    // With no field named there is nothing to inspect: DCOUNT and DCOUNTA count
    // the matching records themselves.
    return { values: [], count: rows.length, omitted: true };
  }

  const column = resolveField(fieldArg, database);
  if (isError(column)) return column;
  return {
    values: rows.map((row) => arrayAt(database, row, column)),
    count: rows.length,
    omitted: false,
  };
}

/** The numeric values among a field's selected cells, with errors propagated. */
function numericValues(values: readonly Scalar[]): number[] | CellError {
  const numbers: number[] = [];
  for (const value of values) {
    if (isError(value)) return value;
    // Text and blanks are skipped rather than coerced, exactly as they are in a
    // SUM over a range.
    if (typeof value === 'number') numbers.push(value);
  }
  return numbers;
}

function databaseSpec(
  name: string,
  summary: string,
  fieldRequired: boolean,
  reduce: (values: Scalar[], recordCount: number, fieldOmitted: boolean) => Value,
): FunctionSpec {
  return {
    name,
    params: [p.array('database'), p.any('field', !fieldRequired), p.array('criteria')],
    summary,
    impl: (args, ctx) => {
      const selected = fieldValues(args, fieldRequired, ctx);
      if (isError(selected)) return selected;
      return reduce(selected.values, selected.count, selected.omitted);
    },
  };
}

/** The shared body of the statistical D functions. */
function numericSpec(
  name: string,
  summary: string,
  reduce: (numbers: number[]) => Value,
): FunctionSpec {
  return databaseSpec(name, summary, true, (values) => {
    const numbers = numericValues(values);
    return isError(numbers) ? numbers : reduce(numbers);
  });
}

function sumOf(numbers: readonly number[]): number {
  let total = 0;
  for (const n of numbers) total = excelAdd(total, n);
  return total;
}

/** The sum of squared deviations, shared by the variance and deviation pair. */
function sumOfSquares(numbers: readonly number[]): number {
  const mean = sumOf(numbers) / numbers.length;
  let total = 0;
  for (const n of numbers) {
    const deviation = n - mean;
    total = excelAdd(total, deviation * deviation);
  }
  return total;
}

const DSUM = numericSpec('DSUM', 'The sum of a field over the matching records.', sumOf);

const DAVERAGE = numericSpec(
  'DAVERAGE',
  'The average of a field over the matching records.',
  (numbers) => (numbers.length === 0 ? CellError.DIV0 : sumOf(numbers) / numbers.length),
);

const DMAX = numericSpec(
  'DMAX',
  'The largest value of a field over the matching records.',
  // An empty selection is zero, not an error: Excel's D aggregates behave like
  // MAX over an empty range.
  (numbers) => (numbers.length === 0 ? 0 : Math.max(...numbers)),
);

const DMIN = numericSpec(
  'DMIN',
  'The smallest value of a field over the matching records.',
  (numbers) => (numbers.length === 0 ? 0 : Math.min(...numbers)),
);

const DPRODUCT = numericSpec(
  'DPRODUCT',
  'The product of a field over the matching records.',
  (numbers) => {
    if (numbers.length === 0) return 0;
    let product = 1;
    for (const n of numbers) product *= n;
    return checkMagnitude(product);
  },
);

const DVAR = numericSpec(
  'DVAR',
  'The sample variance of a field over the matching records.',
  (numbers) =>
    numbers.length < 2 ? CellError.DIV0 : sumOfSquares(numbers) / (numbers.length - 1),
);

const DVARP = numericSpec(
  'DVARP',
  'The population variance of a field over the matching records.',
  (numbers) => (numbers.length === 0 ? CellError.DIV0 : sumOfSquares(numbers) / numbers.length),
);

const DSTDEV = numericSpec(
  'DSTDEV',
  'The sample standard deviation of a field over the matching records.',
  (numbers) =>
    numbers.length < 2
      ? CellError.DIV0
      : Math.sqrt(sumOfSquares(numbers) / (numbers.length - 1)),
);

const DSTDEVP = numericSpec(
  'DSTDEVP',
  'The population standard deviation of a field over the matching records.',
  (numbers) =>
    numbers.length === 0 ? CellError.DIV0 : Math.sqrt(sumOfSquares(numbers) / numbers.length),
);

const DCOUNT = databaseSpec(
  'DCOUNT',
  'The count of numeric cells in a field over the matching records.',
  false,
  (values, recordCount, omitted) =>
    omitted ? recordCount : values.filter((v) => typeof v === 'number').length,
);

const DCOUNTA = databaseSpec(
  'DCOUNTA',
  'The count of non-blank cells in a field over the matching records.',
  false,
  (values, recordCount, omitted) =>
    omitted ? recordCount : values.filter((v) => v !== null && v !== '').length,
);

const DGET = databaseSpec(
  'DGET',
  'The single value of a field in the one matching record.',
  true,
  (values, recordCount) => {
    // Excel distinguishes the two failures: nothing matched is #VALUE!, more
    // than one match is #NUM!.
    if (recordCount === 0) return CellError.VALUE;
    if (recordCount > 1) return CellError.NUM;
    return values[0] ?? null;
  },
);

export const ENGINEERING_FUNCTIONS: readonly FunctionSpec[] = [
  BIN2DEC, BIN2HEX, BIN2OCT,
  DEC2BIN, DEC2HEX, DEC2OCT,
  HEX2BIN, HEX2DEC, HEX2OCT,
  OCT2BIN, OCT2DEC, OCT2HEX,
  BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT,
  COMPLEX, IMABS, IMAGINARY, IMREAL, IMARGUMENT, IMCONJUGATE,
  IMSUM, IMSUB, IMPRODUCT, IMDIV, IMPOWER, IMSQRT, IMEXP,
  IMLN, IMLOG2, IMLOG10,
  IMSIN, IMCOS, IMTAN, IMSINH, IMCOSH, IMSEC, IMCSC, IMCOT,
  DELTA, GESTEP, CONVERT,
  ERF, ERFC, ERF_PRECISE, ERFC_PRECISE,
  BESSELI, BESSELJ, BESSELK, BESSELY,
  DSUM, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT,
  DSTDEV, DSTDEVP, DVAR, DVARP,
];
