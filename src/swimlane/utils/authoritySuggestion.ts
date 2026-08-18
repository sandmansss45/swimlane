import { AuthorityTier } from '../models/IProcessStep';

/**
 * Confirmed rule: job-title suggestions differentiate by authority tier
 * (Execute -> analyst, Endorse -> manager, Approve -> senior/chief).
 *
 * IMPORTANT: this is a keyword heuristic, not a real AI call. An SPFx web
 * part runs entirely in the browser, so it can't safely hold an API key
 * for a real AI suggestion service (anyone could open dev tools and steal
 * it) - a real AI-backed version needs a secure server-side proxy (an
 * Azure Function or Power Automate flow that holds the key), which isn't
 * built yet. This heuristic is a reasonable stand-in until that exists.
 */
export function matchesAuthorityTier(jobTitle: string, tier: AuthorityTier): boolean {
  const normalized = (jobTitle || '').toLowerCase();
  if (tier === 'senior') return /chief|director|head of|vp|vice president/.test(normalized);
  if (tier === 'manager') return /manager|lead/.test(normalized);
  return /analyst|specialist|coordinator|associate/.test(normalized);
}
