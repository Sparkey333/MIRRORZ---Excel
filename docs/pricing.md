# Pricing

Free and fully functional first, sold second. The free tier opens, edits and
saves every format the product supports, and that is a permanent guarantee
enforced in code (`packages/licensing/src/entitlements.ts`, `NEVER_GATED`), not a
marketing position that can be quietly walked back in version 2.

Two rules govern the ladder below:

1. **The one-time price is a real but modest saving**, not a punitive one. If
   buying the perpetual licence costs three years of subscription, it is not an
   option, it is a decoy.
2. **A perpetual licence never stops working.** What expires is the right to
   *future updates*. The build you paid for keeps running, offline, unchecked,
   for as long as the machine boots. Checked against the running build's
   compile-time date, never against the clock.

## The ladder

Prices in USD. The merchant of record displays VAT-inclusive prices where local
law requires it (see *Taking the money*).

| | Free | Personal | Pro | Team (3+ seats) |
|---|---|---|---|---|
| Monthly | - | $6 | $12 | $18 / seat |
| Annual | - | $59 | $119 | $179 / seat |
| **One-time, perpetual** | - | **$99** | **$199** | not offered |
| Update renewal after the first 12 months | - | $39 / yr | $79 / yr | - |
| Devices per licence | unlimited | 2 | 3 | 2 / seat |

### Free - $0, forever, no account, no e-mail address

Everything to do with your own files:

- Open every format we support: .xlsx, .xlsm, .xltx, .xlsb, .xls, .csv, encrypted
  workbooks, and files containing macros, pivots and charts.
- **Edit and save all of them.** Full grid editing, the complete formula
  function library, styles, number formats, sheets, defined names, data
  validation, conditional formatting, find and replace, persistent undo.
- Save, Save As, print, export CSV, export PDF.
- Unknown parts of a file - macros, pivot caches, chart definitions - are
  preserved byte for byte on save. A free user's file survives a round trip
  intact. This is the single most important promise in the product and it is not
  for sale.
- View VBA source read-only, render existing charts, render cached pivot tables.
- The local scripting runtime, deliberately ungated.

### Personal - $6/mo, $59/yr, or $99 once

Free, plus the things an individual power user upgrades for: chart authoring and
editing, local version history with a private per-file store, and the formula
inspector's root-cause search. Personal use, one person.

### Pro - $12/mo, $119/yr, or $199 once

Personal, plus the expensive engineering: **VBA macro execution**, the **live
pivot engine** (re-pivot, filter, drill down), **semantic diff** and **three-way
merge** between workbooks, and **headless batch conversion** from the command
line. Licensed for commercial use by one person.

### Team - $18/seat/mo or $179/seat/yr

Pro, plus seat issue and reassignment and a support queue. **Deliberately no
perpetual option**: perpetual licences plus seat administration is a tar pit for
a solo developer, and saying so is more honest than pricing it out of reach.

## What the one-time licence actually buys

The words a buyer should read, and the words the software has to honour:

> $199 buys MIRRORZ Sheets 1.x forever, plus every update we ship in the next
> 12 months. After that it keeps working exactly as it is. Renew at $79/yr only
> if you want new versions. There is no check, no server, and nothing to renew
> if you are happy with what you have.

Mechanically: the licence carries `maintenanceExpires`, and the *build* carries a
compile-time release date. A build released inside your coverage window runs at
your full plan forever. A build released after it runs at the free tier and says
so plainly, while your covered build keeps working - so the worst case of letting
maintenance lapse is that you do not install a newer version, never that you lose
what you bought. There is no wall-clock condition on a perpetual licence at all.

## Upgrade pricing for a later major version

- **50% of the then-current one-time price** for any existing perpetual holder,
  with no time limit on when you upgraded from. A 1.x holder in 2031 pays half.
- **Free** if you bought within 90 days of the new major shipping, applied
  automatically, no support ticket.
- **Renewed maintenance covers a major version** that ships inside your window.
  Paying $79/yr and then being charged again for 2.0 released during that year
  would be a bait and switch.
- Subscribers get major versions as part of the subscription. That is the
  subscription's actual advantage and we should say so instead of pretending the
  perpetual buyer is getting a worse deal in some unnamed way.

## Why these numbers

- **One-time ÷ annual = 1.68×** for both paid tiers. Break-even against annual is
  20 months, against monthly 16.5 months. A genuine saving for anyone who intends
  to stay, without cannibalising the subscription.
- The comparable band recorded in this repo's research pass is roughly
  **1.7×-2.8×** (TablePlus at the friendly end, SoftMaker Office at the other),
  so we sit at the customer-friendly end deliberately.
- **Annual saves ~18% on monthly** - about two months free. Legible without being
  a gimmick.
- Against the anchors a buyer already holds: Personal at $59/yr undercuts
  Microsoft 365 Personal (~$99.99/yr as recorded) by around 40%, and $99 one-time
  undercuts Office Home 2024 ($149.99 as recorded) by a third.
- 50% education and non-profit discount, applied on request with no proof
  theatre beyond a plausible address.

## Why the paywall sits where it does

Two independent teams have already market-tested where the willingness to pay is
in this exact category: Univer (Apache-2.0, VC-backed) put charts, pivot tables
and collaboration in its commercial tier, and SheetJS put styling, charts, pivots
and formula evaluation in SheetJS Pro. We follow the shape of that finding but
move the line further towards the user - **formula evaluation and file I/O are
free here, and both of those companies charge for at least one of them** - and we
add the one thing nobody else can offer: VBA execution.

Every gated capability in `entitlements.ts` carries a written reason. The rule we
hold ourselves to is that if the reason cannot be written honestly, the
capability does not get gated.

## Taking the money

A solo developer selling worldwide has a tax problem long before they have a
revenue problem: the EU has **no registration threshold** for non-EU sellers of
digital services, so the first €1 of EU revenue creates an obligation. A merchant
of record (MoR) becomes the seller of record, calculates, collects and **remits**
VAT/GST/sales tax in its own name, and absorbs the compliance. That is what the
extra percent buys, and for one person it is cheap.

| | Fee (as recorded) | Merchant of record? | Notes |
|---|---|---|---|
| **Polar** | ~4-5% + 40-50¢ | Yes | Native licence-key issuing with per-key activation limits, and **unauthenticated** validate/activate endpoints, so no API secret needs to ship in the binary. Their own codebase is Apache-2.0. Cheapest MoR recorded. |
| **Paddle** | 5% + $0.50 | Yes | The only one of these that handles **EU/UK B2B VAT-ID reverse charge** properly. Vets sellers, so apply early. |
| **Lemon Squeezy** | ~5% + 50¢ | Yes | Acquired by Stripe and being folded into Stripe Managed Payments through 2026. **Do not start here** - the migration risk is the whole objection. |
| **Stripe (direct)** | 2.9% + 30¢ headline | **No** | Stripe Tax *calculates* but does not *remit*; you register and file everywhere yourself. All-in for a global solo seller lands nearer 7.2%-11.7% once international cards, FX and disputes are counted. Only sensible after incorporation, with an accountant. |
| Gumroad | ~12.9% + $0.80 all-in | Yes | Acceptable to take the first legitimate $99 with zero paperwork during a soft launch. Never the steady state. |

**Recommendation:** start on **Polar**, add **Paddle** when meaningful EU/UK B2B
appears. Revisit only after incorporation.

Two things that are true regardless of processor: the MoR sees the customer's
e-mail and country because tax law requires it, and **the application never
does** - the licence is minted on our machine and typed into theirs, and nothing
in the binary reports back.

## Refunds

- **30 days, no questions, no forms, no retention flow.** Reply to the receipt.
- Refunds are processed by the merchant of record, so they land on the original
  card without us handling card data.
- **We do not revoke on refund.** With offline verification we could not enforce
  a revocation anyway, and a licence-check that phones home is precisely the
  thing this product refuses to build. Someone who refunds and keeps using it has
  taken $99 from a solo developer, and we would rather carry that loss than build
  the surveillance needed to prevent it. Say this out loud on the pricing page:
  it is a better advertisement than any feature.
- Annual subscriptions cancelled mid-term keep running to the end of the term.
- Cancellation is one button in the customer portal. No phone call, no "are you
  sure", no dark pattern.

## What we will never do

- **No advertising.** Ever, in any tier, including free.
- **No telemetry.** No usage analytics, no crash pings without an explicit
  per-incident prompt, no "anonymous" counters. The application makes no network
  request unless you ask it to. Licence verification is a local signature check.
- **No account, no cloud requirement.** You can install and use this on a machine
  that has never had a network connection, and it will work identically.
- **No holding files hostage.** An expired, tampered, forged or absent licence
  degrades to the free tier and nothing else. It never disables saving, never
  watermarks a document, never makes a file read-only, never refuses to export,
  and never deletes anything. Your files are yours.
- **No selling or sharing anything about you**, because we do not collect it.
- **No paid-tier-only file format.** Everything we can write, every tier can
  write.
- **No expiry on a perpetual licence.** If we ever ship a build that stops a paid
  perpetual licence from running, that is a bug of the highest severity.

## Sources and verification status

Every figure above that concerns a third party came from this repository's
earlier research pass and is recorded here so it can be re-checked. **The
research environment for this pass had no outbound network access, so none of
these were re-verified first-hand.** Treat the whole table as needing
confirmation against the primary source before any of it appears on a public
pricing page, in an investor deck or in an advertisement - the product plan
already mandates this mitigation for exactly this reason.

| Claim | Where to verify |
|---|---|
| Microsoft 365 Personal ~$99.99/yr; Office Home 2024 $149.99 one-time | microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products |
| TablePlus ~$99 + ~$59/yr renewal (≈1.7× ratio) | tableplus.com/pricing |
| SoftMaker Office ~$149.95 perpetual vs ~$49.90/yr (≈2.8×) | softmaker.com/en/softmaker-office-shop |
| Polar ~4-5% + 40-50¢, MoR, unauthenticated licence validation | polar.sh/pricing and their licensing docs |
| Paddle 5% + $0.50, MoR, reverse-charge handling | paddle.com/pricing |
| Lemon Squeezy ~5% + 50¢ and the Stripe acquisition/migration | lemonsqueezy.com/pricing and Stripe's acquisition announcement |
| Stripe 2.9% + 30¢, Stripe Tax calculates but does not remit | stripe.com/pricing and stripe.com/tax |
| Gumroad ~12.9% + $0.80 all-in | gumroad.com/pricing |
| EU: no registration threshold for non-EU digital-service sellers | VAT One Stop Shop guidance, ec.europa.eu |
| Univer and SheetJS commercial-tier feature splits | univer.ai/pricing, sheetjs.com/pro |

Currency, tax treatment and every price above are subject to change by their
owners; a figure that is a year old is a guess.
