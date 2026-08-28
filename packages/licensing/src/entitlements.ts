/**
 * The feature gate: which plan may do what.
 *
 * The governing rule, and it is not negotiable anywhere in this file: OPENING,
 * EDITING AND SAVING EVERY SUPPORTED FORMAT IS NEVER GATED. Not by plan, not by
 * trial state, not by a licence that failed to verify, not by a clock that went
 * backwards. A person's spreadsheets are their own work and their own property,
 * and a program that holds them behind a payment is a program that should not be
 * trusted with them in the first place. {@link NEVER_GATED} is unioned into every
 * plan's capability set unconditionally, so no future edit to the tables below
 * can take those away by accident.
 *
 * What is left to sell is the work that is genuinely expensive to build and that
 * a professional user gets paid for using: executing VBA, live pivoting, semantic
 * diff and three-way merge, formula root-cause analysis, and headless batch
 * conversion. Each gated capability below carries the reason it is gated. If a
 * reason cannot be written honestly, the capability does not get gated.
 *
 * Nothing here calls out to a network, and the gate is a capability check on a
 * value computed locally - not a licence server, not a nag timer.
 */

import type { LicenseAssessment } from './license.js';
import type { PlanId } from './plans.js';
import { PLANS, betterPlan } from './plans.js';
import type { TrialState } from './trial.js';

export type Capability =
  // --- never gated ---------------------------------------------------------
  | 'file.open'
  | 'file.save'
  | 'file.saveAs'
  | 'file.print'
  | 'export.csv'
  | 'export.pdf'
  | 'edit.cells'
  | 'edit.structure'
  | 'format.cells'
  | 'formula.calculate'
  | 'undo.persistent'
  | 'find.replace'
  | 'clipboard.interop'
  | 'macro.preserve'
  | 'pivot.preserve'
  | 'chart.render'
  | 'automation.scripts'
  // --- sold -----------------------------------------------------------------
  | 'chart.author'
  | 'history.local'
  | 'macro.execute'
  | 'pivot.interactive'
  | 'diff.semantic'
  | 'merge.threeWay'
  | 'formula.inspector'
  | 'batch.cli'
  | 'use.commercial'
  | 'seats.manage'
  | 'support.priority';

/**
 * The floor. Every plan, including free and including a tampered licence, gets
 * all of these. Read the list as the promise it is.
 */
export const NEVER_GATED: readonly Capability[] = [
  'file.open', // every format we can read, including .xls, .xlsb and encrypted workbooks
  'file.save', // every format we can write, with unknown parts preserved byte for byte
  'file.saveAs',
  'file.print',
  'export.csv',
  'export.pdf', // printing to paper is free, so printing to a file is too
  'edit.cells',
  'edit.structure', // insert, delete, move rows, columns and sheets
  'format.cells',
  'formula.calculate', // the whole function library, not a subset
  'undo.persistent',
  'find.replace',
  'clipboard.interop',
  'macro.preserve', // we round-trip VBA byte for byte even though we will not run it
  'pivot.preserve', // cached pivot tables render and survive a save
  'chart.render',
  'automation.scripts', // the local scripting runtime, deliberately ungated
];

export interface GateReason {
  readonly capability: Capability;
  /** Lowest plan that includes it. */
  readonly plan: PlanId;
  /** Why this one is sold rather than given away. Honest or it is not gated. */
  readonly why: string;
}

/** The complete list of gated capabilities. If it is not here, it is free. */
export const PREMIUM: readonly GateReason[] = [
  {
    capability: 'chart.author',
    plan: 'personal',
    why: 'Authoring and editing charts is a large, separate feature surface; reading and rendering existing charts stays free so no file ever looks broken.',
  },
  {
    capability: 'history.local',
    plan: 'personal',
    why: 'Local version history keeps a private store of every save. It is the feature people upgrade for and it costs real disk and real engineering.',
  },
  {
    capability: 'macro.execute',
    plan: 'pro',
    why: 'A VBA interpreter is the single most expensive thing in this product. Extraction, preservation and read-only source viewing are free; running the code is not.',
  },
  {
    capability: 'pivot.interactive',
    plan: 'pro',
    why: 'The live pivot engine (re-pivot, drill down, filter) is a database in miniature. Existing pivots still render and round-trip for free.',
  },
  {
    capability: 'diff.semantic',
    plan: 'pro',
    why: 'Structural diff between two workbooks is professional review work, and the people doing it are paid for it.',
  },
  {
    capability: 'merge.threeWay',
    plan: 'pro',
    why: 'Three-way merge of two edited copies against a common ancestor. Same reason as diff.',
  },
  {
    capability: 'formula.inspector',
    plan: 'personal',
    why: 'Root-cause search across a dependency graph to explain one wrong number. Ordinary formula evaluation, error values and the error explainer stay free, because a file that shows the wrong number without saying why is a broken file.',
  },
  {
    capability: 'batch.cli',
    plan: 'pro',
    why: 'Headless batch conversion is how the product gets used on a server, which is a commercial use with commercial value.',
  },
  {
    capability: 'use.commercial',
    plan: 'pro',
    why: 'A licence term, not a lock: nothing in the software checks it, and nothing ever will. Free and Personal are for personal work; businesses buy Pro.',
  },
  {
    capability: 'seats.manage',
    plan: 'team',
    why: 'Issuing and reassigning seats across a team.',
  },
  {
    capability: 'support.priority',
    plan: 'team',
    why: 'A queue position, paid for with money rather than with software restrictions.',
  },
];

const PREMIUM_BY_CAPABILITY = new Map<Capability, GateReason>(PREMIUM.map((gate) => [gate.capability, gate]));

function planIncludes(plan: PlanId, gate: GateReason): boolean {
  return PLANS.indexOf(plan) >= PLANS.indexOf(gate.plan);
}

/**
 * Capabilities granted by a plan, plus any per-licence extras. Always starts
 * from {@link NEVER_GATED}, which is the structural guarantee that the free tier
 * cannot be hollowed out by an edit to the tables above.
 */
export function capabilitiesFor(plan: PlanId, extraFeatures: readonly string[] = []): ReadonlySet<Capability> {
  const granted = new Set<Capability>(NEVER_GATED);
  for (const gate of PREMIUM) {
    if (planIncludes(plan, gate)) granted.add(gate.capability);
  }
  // Per-licence extras let us honour a one-off deal or thank a beta tester
  // without inventing a plan. Unknown names are ignored, never fatal.
  for (const feature of extraFeatures) {
    if (PREMIUM_BY_CAPABILITY.has(feature as Capability)) granted.add(feature as Capability);
  }
  return granted;
}

export const PLAN_CAPABILITIES: Readonly<Record<PlanId, ReadonlySet<Capability>>> = Object.freeze(
  Object.fromEntries(PLANS.map((plan) => [plan, capabilitiesFor(plan)])) as Record<PlanId, ReadonlySet<Capability>>,
);

export interface Entitlements {
  readonly plan: PlanId;
  /** Where the plan came from, for the status line and for tests. */
  readonly source: 'license' | 'trial' | 'free';
  readonly capabilities: ReadonlySet<Capability>;
  readonly seats: number;
  readonly explanation: string;
  can(capability: Capability): boolean;
  /** Why a capability is unavailable, or null when it is available. */
  gateFor(capability: Capability): GateReason | null;
}

function build(plan: PlanId, source: Entitlements['source'], seats: number, explanation: string, extras: readonly string[]): Entitlements {
  const capabilities = capabilitiesFor(plan, extras);
  return {
    plan,
    source,
    capabilities,
    seats,
    explanation,
    can: (capability) => capabilities.has(capability),
    gateFor: (capability) => (capabilities.has(capability) ? null : PREMIUM_BY_CAPABILITY.get(capability) ?? null),
  };
}

/**
 * Resolve what the running application may do, from a licence assessment and an
 * optional trial. Whichever grant is higher wins, so installing a Personal
 * licence during a Pro trial never costs the user anything mid-trial.
 */
export function entitlementsFor(assessment: LicenseAssessment, trial?: TrialState | null): Entitlements {
  if (trial && trial.active) {
    if (assessment.plan === 'free') return build(trial.plan, 'trial', 1, trial.explanation, []);

    const plan = betterPlan(assessment.plan, trial.plan);
    const fromLicense = plan === assessment.plan;
    return build(
      plan,
      fromLicense ? 'license' : 'trial',
      assessment.seats || 1,
      fromLicense ? assessment.explanation : trial.explanation,
      assessment.features,
    );
  }

  if (assessment.plan === 'free') {
    return build('free', 'free', 0, assessment.explanation, []);
  }

  return build(assessment.plan, 'license', assessment.seats, assessment.explanation, assessment.features);
}

/** The free tier as a standalone value, for the boot path before anything loads. */
export function freeEntitlements(): Entitlements {
  return build(
    'free',
    'free',
    0,
    'Free tier. Opening, editing and saving every supported format are never restricted.',
    [],
  );
}
