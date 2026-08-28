# Where the LibreOffice oracle disagrees with Excel

Our fixture pipeline recalculates workbooks with LibreOffice headless and treats
the cached results as expected values. That gives us 137+ verified formula cases
for free, and it is right almost everywhere - but it is not Excel, and a handful
of documented divergences must not be baked into our tests as if they were
Excel's behaviour.

Known divergences, to be re-verified against real Excel before we rely on them:

| Case | LibreOffice | Excel | Notes |
| --- | --- | --- | --- |
| `=TRUE>1` | `FALSE` | `TRUE` | Excel ranks types for comparison: number < text < FALSE < TRUE, so any boolean is greater than any number. LibreOffice coerces `TRUE` to `1` and compares numerically. We follow **Excel**. |
| `=0^0` and `=POWER(0,0)` | `1` | `#NUM!` | Nought to the nought is a convention rather than a fact, and the two products chose differently. Excel refuses it from both the operator and the function; LibreOffice answers `1` from both. We follow **Excel**. Measured, not assumed: a probe workbook recalculated by the same pipeline the fixtures use. |
| `=0^-1` and `=POWER(0,-1)` | `#NUM!` | `#DIV/0!` | Both refuse it; they disagree about which refusal it is. Excel calls it what it is - a division by zero - and LibreOffice reports a domain error. We follow **Excel**. |
| `=XLOOKUP(...)`, `=XMATCH(...)` | `#NAME?` | value | LibreOffice 24.2 does not implement these. Not a divergence so much as an oracle **gap**: our engine computes the correct answers where the oracle could not. The engine test carries the documented expected values (`XLOOKUP` -> `121000`, `XMATCH` -> `8`) and verifies against those, so being ahead of the oracle is still checked rather than waved through. |

Cases the oracle **did** confirm, and which we rely on:

- `=2^3^2` is `64`. Exponentiation is left-associative in Excel, unlike normal
  mathematical convention and unlike most parser-generator defaults.
- `=-2^2` is `4`. Unary minus binds tighter than `^`.
- `=2^2%` is `1.0139...`, i.e. `2^(2%)`. Postfix `%` binds tighter than `^`.
- `=0.1+0.2` displays as `0.3` and `=0.1+0.2=0.3` is `TRUE`.
- `=(0.5-0.4-0.1)` is exactly `0`, and comparing it to `0` is `TRUE`.
- `=1/3` is `0.333333333333333` - fifteen significant digits, not seventeen.
- `=9007199254740993` stores as `9007199254740990`: numeric literals are
  truncated to fifteen significant digits on entry.
- `=ROUND(2.5,0)` is `3` and `=ROUND(-2.5,0)` is `-3`: half away from zero, not
  banker's rounding.
- `="a">1` is `TRUE`: text sorts above numbers.
- A blank cell coerces to `0` in arithmetic and `""` in concatenation, and
  compares equal to both.
- Space is the intersection operator and a parenthesised comma list is a union:
  `SUM(F40:H42 G40:G45)` and `SUM((F40:F41,F42:F43))` both evaluate.
