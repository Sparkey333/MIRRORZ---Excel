# Naming, trademark and fonts

Not legal advice. This is the engineering-side statement of what we may ship,
written so that no one has to guess in a pull request. Have counsel confirm it
before the first paid download.

## 1. The name is a problem right now

The repository is called `MIRRORZ---Excel`. **"Excel" cannot appear in the name
of the product we sell**, and that is the single highest-severity legal item in
this project.

Why it is worse than the usual "you may not use someone's trademark" caution:

- **Excel** and **Microsoft** are registered trademarks of Microsoft
  Corporation. Trademark infringement turns on *likelihood of confusion*, and
  every factor that matters points the wrong way here: identical goods (a
  spreadsheet application), identical customers, identical channels, and a mark
  used in the product name rather than in a sentence about it.
- Microsoft's own Trademark and Brand Guidelines prohibit using their marks in
  "the name of your business, product, service, app, domain name, social media
  account" - and the whole point of a guidelines page is that Microsoft points
  at it when it sends a letter.
- We are a **directly competing product**. Nominative fair use protects
  describing what a product does; it does not stretch to naming yourself after
  the competitor you are displacing.
- App stores enforce this independently. A Microsoft Store or Mac App Store
  listing with "Excel" in the title, subtitle or keyword field gets rejected or
  pulled, and by then the URL, the reviews and the install base are attached to
  the name.

### The recommendation

**Ship as "MIRRORZ Sheets".** The code already uses it in places; make it the
only name that exists outside the repository.

That name must be the one in: the product name and window title, the About box,
the installer and app-bundle name, the macOS bundle identifier and Windows
publisher/product strings, the code-signing certificate subject, the domain, the
social handles, every store listing, and the update feed. The repository name
should be changed too - `mirrorz-sheets` - because repository URLs end up in
release notes, `package.json` files, crash reports and third-party licence
documents, which are all shipped artefacts in every way that matters.

Two things to do before spending money on the MIRRORZ brand itself: run a
clearance search (USPTO and EUIPO, classes 9 and 42, plus a plain web search for
existing software called "Mirror(s)/Mirrorz"), and secure the domain and handles
in one pass. A rename after launch costs the install base; a rename now costs an
afternoon.

## 2. Safe wording

The governing rule: **use "Excel" only as an adjective, immediately before a
generic noun, in a sentence that describes compatibility.** Never as a noun, never
as our own name, never in a way that implies endorsement, affiliation, licensing
or certification.

### Say this

- "Opens and edits .xlsx files."
- "Compatible with Excel formats."
- "Reads and writes .xlsx, .xlsm, .xlsb, .xls and .csv."
- "Opens files created in Microsoft Excel."
- "Excel-compatible formulas" / "Excel file formats."
- "Formulas behave the way Microsoft Excel calculates them."
- In a file dialog filter: "Excel Workbook (*.xlsx)" - this names the format,
  which is the textbook case for nominative use, and every other office suite
  does it. Keep the extension in the string.

### Never say this

- **"MIRRORZ Excel"**, or any product, edition, tier or SKU name containing
  "Excel" or "Microsoft" or "Office".
- **"Excel for Mac"**, "Excel for Linux", "the Excel replacement", "Excel Pro" -
  anything that reads as a *version of* Excel rather than a *different program*.
- "Official", "certified", "approved", "endorsed", "partner", "powered by
  Microsoft", "Microsoft-compatible" (that last one implies a certification
  programme that we are not in).
- Anything using Microsoft's logos, the Excel icon, the Excel green, or an icon
  designed to be mistaken for it at 32 pixels. Our icon must not be a green
  document with an X.
- "Excel" as the first word of a headline, a store listing title, a page title,
  or a domain.

### Attribution and disclaimer

On the website footer, in the About box, and in the README:

> Microsoft and Excel are registered trademarks of Microsoft Corporation.
> MIRRORZ Sheets is not affiliated with, endorsed by, or sponsored by Microsoft.
> "Excel" is used only to describe file-format compatibility.

Use the marks in plain text, in the same typeface as the surrounding copy. Do not
style them, do not add ™/® to *our* uses beyond a first-use attribution, and do
not put them anywhere they could read as a co-brand.

### Enforcement, not convention

The build already plans a grep gate over marketing copy and the app bundle
asserting that "Excel", "Microsoft" and "Office" appear only adjectivally before
a generic noun. Keep it a hard CI failure. Naming discipline decays the moment it
depends on someone remembering.

## 3. Fonts: Calibri and Cambria cannot be shipped

**Calibri** and **Cambria** are proprietary font software commissioned by
Microsoft and licensed with Windows and Office. Redistributing either one inside
our installer, our app bundle, a web font, or an exported PDF's embedded subset
is copyright infringement of the font software, independently of anything to do
with trademarks. There is no version of this that is fine because "the file
already contained it".

Two distinctions that keep coming up:

- **Using an installed font is fine. Redistributing it is never fine.** If the
  user's machine has Calibri because they have Office, we may render with it.
- **Metrics are not glyphs.** A metric-compatible substitute has the same advance
  widths, so column autofit, row heights, wrap points and pagination match
  Excel's. The letterforms differ slightly. That is exactly the trade we want,
  because layout fidelity is what users notice and glyph identity is not.

### Bundle these instead

| Microsoft font | Metric-compatible substitute | Licence |
|---|---|---|
| Calibri | **Carlito** (Łukasz Dziedzic) | SIL Open Font License 1.1 |
| Cambria | **Caladea** (Huerta Tipográfica) | Apache-2.0 |
| Arial | Arimo | Apache-2.0 |
| Times New Roman | Tinos | Apache-2.0 |
| Courier New | Cousine | Apache-2.0 |

Obligations that attach to shipping them, all cheap and all easy to forget:

- Ship each font's licence text, keep the copyright notices, and list the fonts
  in the build-time third-party licence document surfaced in Help → Licences and
  in the installer.
- SIL OFL: do not sell the font on its own, and do not reuse a **Reserved Font
  Name** if we ever modify or subset-and-rename a face.
- **Do not present Carlito to the user as "Calibri."** Substitute silently in
  layout, but show the substitution honestly in the font list and in any "fonts
  missing" notice. Relabelling someone else's font with Microsoft's trademarked
  font name would recreate the trademark problem inside the font picker.

### Aptos is an unsolved and compounding problem

Microsoft replaced Calibri with **Aptos** as the Office default in 2023-24. Aptos
is proprietary, ships only with Microsoft 365, and - unlike Calibri, which got
Carlito - **has no free metric-compatible clone**. Every quarter, a larger share
of incoming workbooks specifies a font we cannot ship and cannot match, so on
macOS and Linux we will substitute at different widths and produce wrong autofit,
wrong row heights, wrong wrap points and wrong pagination against Excel.

Mitigation, in order of cost: use system-installed Aptos on Windows when it is
present; maintain an explicit substitution table with cached real metrics for the
platforms where it is absent; publish the substitution in the compatibility
matrix so it is a documented limitation rather than a surprise; and seriously
consider co-funding a metric-compatible substitute, since this is now an unfilled
gap shared by every non-Microsoft office suite and the cost is plausibly
shareable.

## 4. The one-line version

Ship as **MIRRORZ Sheets**. Say **"opens and edits .xlsx files"** and
**"compatible with Excel formats"**. Never say **"Excel for Mac"** or anything
implying affiliation. Bundle **Carlito and Caladea**, never Calibri or Cambria,
and be honest in the UI about which one the user is looking at.
