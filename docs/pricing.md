# Pricing

Researched against comparable prosumer desktop productivity apps that offer both
subscription and perpetual options. Two rules govern the whole ladder: the
one-time price is a real but modest saving against subscribing rather than a
punitive one, and a perpetual licence never stops working.

## The ladder

Prices in USD, charged in USD, with the merchant of record handling
VAT-inclusive display.

| Tier | Monthly | Annual | One-time | Update renewal | Devices |
|---|---|---|---|---|---|
| Reader | free | — | — | — | unlimited |
| Personal | $6/mo | $59/yr | **$99** | $39/yr | 2 |
| Pro | $12/mo | $119/yr | **$199** | $79/yr | 3 |
| Team (3+ seats, commercial) | $18/seat/mo | $179/seat/yr | subscription only | — | 2/seat |

## What is in each tier

 Reader (free, forever, no account): open, view, print, export CSV, read-only, every format we support — including .xls, .xlsb, encrypted workbooks and files containing macros/pivots/charts, all preserved on any save-as-copy. This is the growth engine, not charity: it gets people to register the file association, and the file association is the habit that sells the upgrade. It also makes the strongest possible statement of the product thesis — we will never destroy your file, even for free. Personal: full editing, all read/write formats except macro execution, the standard formula engine, the sheet explorer, regex find/replace, unlimited persistent undo, non-commercial use. Pro: VBA macro execution, the large-file engine, pivot tables, local version history + semantic diff + merge, the formula inspector with root-cause search, the local scripting runtime, PDF export, commercial use for one person. Team: adds seat management and priority support; deliberately NO perpetual option, because perpetual plus seat management is an administrative tar pit.

## The one-time offer, in the words the customer should read

 "$99 buys MIRRORZ Sheets 1.x forever, plus every update we ship in the next 12 months. After that it keeps working exactly as it is — renew at $39/yr only if you want new versions. Upgrade to 2.0 at 50% off." That is a fair, modest, honest deal: the software genuinely never expires, it never phones home, and an old build runs offline forever with no check (enforced by comparing the BUILD's compile-time-baked date against the license's updatesUntil, never by expiring the license).

## Why these numbers

 One-time ÷ annual = 1.68× for both tiers, so break-even against annual is 20 months and against monthly is 16.5 months — a real discount for anyone who intends to stay, without cannibalising the subscription. The market band for this ratio is 1.7–2.8× (TablePlus ~1.7× at $99 + $59/yr renewal; SoftMaker ~2.8× at $149.95 perpetual vs $49.90/yr), so we sit at the customer-friendly end. Annual saves 18% on monthly (roughly two months free), which is legible without being a gimmick. Against the anchors a buyer actually holds in their head: Personal at $59/yr undercuts Microsoft 365 Personal (~$99.99/yr) by 40%, and $99 one-time undercuts Office Home 2024 ($149.99) by a third — we win both comparisons in the three seconds a buyer spends making them.

## Why the paywall sits where it does

 Univer — Apache-2.0, VC-backed — put import/export, printing, charts, pivot tables, sparklines and collaboration in its COMMERCIAL tier, and SheetJS put styling, charts, pivots and formula evaluation in SheetJS Pro. Two independent teams have already market-tested where the willingness to pay is. We follow it, with one addition nobody else can offer: VBA execution, which is the single biggest reason enterprises cannot leave Excel and which no cross-platform non-Microsoft spreadsheet runs correctly.

## Operations: taking the money

 Merchant of record: Polar (~4–5% + 40–50¢, native license keys, per-key activation limits, and UNAUTHENTICATED validate/activate endpoints so no API secret ships in the binary; their own codebase is Apache-2.0). Move to or add Paddle (5% + $0.50 flat) when meaningful EU/UK B2B appears, because Paddle is the only one of these that handles VAT-ID reverse charge properly — and apply early, they vet sellers. Gumroad (~12.9% + $0.80 all-in) is acceptable for a soft launch to take the first legitimate $99 with zero tax paperwork, never as the steady state. Do NOT use Stripe direct until incorporated with an accountant: Stripe is NOT a merchant of record, Stripe Tax calculates but does not remit, and the EU has no registration threshold for non-EU sellers of digital services — all-in cost for a global solo seller runs 7.2%–11.7% once international card fees, FX and disputes are counted, versus a flat 4–5% at an MoR who also absorbs the compliance. Do not start on Lemon Squeezy; Stripe acquired it and is folding it into Stripe Managed Payments through 2026.

## Policy

 50% education and non-profit discount. No-questions 30-day refund — with an offline-verifiable license we cannot enforce a revocation anyway, so make it a selling point rather than a fight. Do NOT also offer a JetBrains-style perpetual fallback on the subscription: one perpetual path is legible, two invite arbitrage and remove the reason to ever buy the one-time SKU.

## Cost to get there

Total pre-revenue outlay to a fully signed, notarized, auto-updating, license-gated cross-platform product: $99/yr (Apple) + ~$120/yr (Azure Artifact Signing Basic) ≈ $220/year, plus GitHub Actions minutes (free on public repos).

## What we will never do

- No advertising, ever.
- No telemetry. The application makes no network request unless the user asks it
  to, and licence verification is entirely offline.
- No cloud requirement and no account to open a file.
- No holding files hostage. If a licence expires, is tampered with, or is absent,
  the application degrades to the free reader tier - it never locks the user out
  of their own data, and it never refuses to save a copy.
- No dark patterns on cancellation, and a no-questions refund inside 30 days.
  With an offline-verifiable licence a revocation is unenforceable anyway, so it
  is offered as a selling point rather than fought over.

