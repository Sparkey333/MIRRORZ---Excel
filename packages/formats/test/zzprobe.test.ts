import { describe, it } from 'vitest';
import { createRequire } from 'node:module';
import { format } from '../src/numfmt.js';

const require = createRequire('/home/user/MIRRORZ---Excel/');
const SSF = require('/tmp/claude-0/-home-user-MIRRORZ---Excel/f976ef7a-d2e9-5382-a26f-e144d39bb79c/scratchpad/ssf2.js');

const codes = [
  'General','0','0.00','#,##0','#,##0.00','0%','0.00%','0.00E+00','# ?/?','# ??/??',
  'mm-dd-yy','d-mmm-yy','d-mmm','mmm-yy','h:mm AM/PM','h:mm:ss AM/PM','h:mm','h:mm:ss',
  'm/d/yy h:mm','#,##0 ;(#,##0)','#,##0 ;[Red](#,##0)','#,##0.00;(#,##0.00)','mm:ss',
  '[h]:mm:ss','mmss.0','##0.0E+0','@','0%%','0.0%%','0.00%;[Red]-0.00%',
  '#,##0.00_);[Red](#,##0.00)','_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  '000-0000','0-0','#.##','0.','0.??','?.?','# ?/16','?/?','# ???/???','0.0E-0','#0.0E+0',
  '0;;','0.00;;"zed"','"a";"b";"c"','[>100]"big";[<=100]"small"','[Red]0.00','[Blue]#,##0',
  '0.00;@','mm-dd-yy;@','yyyy-mm-dd','dddd, mmmm d, yyyy','[$USD-409]#,##0.00','[$-409]yyyy',
  '[$-en-US]yyyy-mm-dd','yyyy"年"m"月"','h"x"mm','m:[s]','[h]:mm','[mm]:ss','[ss].00',
  '0.0,,','#,##0,','#,','0,000','0.00000000000000000','##0.0E+0','0.0E+0',
  'h:mm:ss.0','hh:mm:ss.00','\\y\\ymmdd','0"abc"','_(0.0_)','0*x0','0.00;-0.00;"z";"lit"',
  'General" units"','[<=100]0.00;@','[>100]0.00','0;[Red]0','#,##0;(#,##0);"-"',
];
const values = [0, 1, -1, 0.5, -0.5, 5, 12, 1234.5678, -1234.5678, 12345, 0.015, 1.5,
  0.7, 5.25, 0.26, 100.5, 1e10, 1e11, 1/3, 0.1+0.2, 45000, 45000.5678, 0.5678, 60, 61,
  1e-10, 9.99, 255, 12200000, 5551234, 0.0001, 100000, 1.9999, 2.675, 1.005, 0.999,
];

describe('probe', () => {
  it('diffs against SSF', () => {
    const rows: string[] = [];
    for (const c of codes) {
      for (const v of values) {
        let mine: string, theirs: string;
        try { mine = format(v as never, c).text; } catch (e) { mine = 'THROW:' + (e as Error).message; }
        try { theirs = SSF.format(c, v); } catch (e) { theirs = 'THROW:' + (e as Error).message; }
        if (mine !== theirs) rows.push(`${JSON.stringify(c)}\t${v}\tMINE=${JSON.stringify(mine)}\tSSF=${JSON.stringify(theirs)}`);
      }
    }
    console.log('DIFFS: ' + rows.length);
    console.log(rows.join('\n'));
  });
});
