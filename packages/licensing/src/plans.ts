/**
 * The plan vocabulary, kept in its own module so the licence codec, the
 * entitlement table and the trial can all name plans without importing each
 * other. Plan codes are wire values: once a licence has been issued with a code
 * it can never be reused for anything else, so append, never reorder.
 */

export const PLANS = ['free', 'personal', 'pro', 'team'] as const;

export type PlanId = (typeof PLANS)[number];

/** Wire codes for the compact licence payload. Append only. */
export const PLAN_CODES: Readonly<Record<PlanId, number>> = {
  free: 0,
  personal: 1,
  pro: 2,
  team: 3,
};

const BY_CODE = new Map<number, PlanId>(PLANS.map((plan) => [PLAN_CODES[plan], plan]));

export function planFromCode(code: number): PlanId | undefined {
  return BY_CODE.get(code);
}

export function isPlanId(value: unknown): value is PlanId {
  return typeof value === 'string' && (PLANS as readonly string[]).includes(value);
}

/**
 * Ordering used when two grants overlap (a licence and a running trial, say).
 * The higher rank wins, so a user is never dropped to a lower tier by owning
 * more than one entitlement.
 */
export function planRank(plan: PlanId): number {
  return PLANS.indexOf(plan);
}

export function betterPlan(a: PlanId, b: PlanId): PlanId {
  return planRank(a) >= planRank(b) ? a : b;
}
