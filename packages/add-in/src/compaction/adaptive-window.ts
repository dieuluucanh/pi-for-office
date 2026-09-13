/**
 * Session-scoped adaptive effective context window.
 *
 * When a run fails with a stream-end/overflow classification while the
 * request-facing estimate is a meaningful fraction of the claimed window (e.g.
 * an OpenCode Go contributor tier with a lower real cap than the catalog says),
 * record the failure estimate and use a reduced effective window for
 * subsequent compaction triggers and guards.
 *
 * The tracker is deliberately conservative: it only downgrades on
 * overflow-like failures at ≥50% of the claimed window, applies a 0.9× margin,
 * and resets when the model changes. It is purely advisory for the guards —
 * never stored in the transcript.
 */

export interface AdaptiveWindowState {
 /** Limited (reduced) effective window, or the claimed window when unused. */
 effectiveWindow: number;
 /** Whether a downgrade is currently active. */
 downgraded: boolean;
 /** The observed failure estimate that triggered the downgrade. */
 lastFailureEstimate: number;
}

export function createAdaptiveWindow(claimedWindow: number): {
 get: () => AdaptiveWindowState;
 /** Reset on model switch or explicit override. */
 reset: (nextClaimedWindow: number) => void;
 /**
  * Consider a failure that ended with an overflow/stream-end classification.
  * `effectiveTokens` is the request-facing estimate at failure time.
  */
 observeFailure: (args: {
  failureIsOverflowLike: boolean;
  effectiveTokens: number;
  contextWindow: number;
 }) => void;
} {
 let effectiveWindow = normalize(claimedWindow);
 let downgraded = false;
 let lastFailureEstimate = 0;

 const setEffectiveWindow = (next: number): void => {
  effectiveWindow = normalize(next);
 };

 return {
  get: () => ({ effectiveWindow, downgraded, lastFailureEstimate }),

  reset: (nextClaimedWindow: number) => {
   setEffectiveWindow(nextClaimedWindow);
   downgraded = false;
   lastFailureEstimate = 0;
  },

  observeFailure: ({
   failureIsOverflowLike,
   effectiveTokens,
   contextWindow,
  }) => {
   const claimed = normalize(contextWindow);
   if (claimed <= 0) return;

   // Only downgrade on overflow-like failures with a meaningful estimate.
   if (!failureIsOverflowLike) return;
   if (effectiveTokens < claimed * 0.5) return;

   const suggested = Math.floor(effectiveTokens * 0.9);
   if (suggested < effectiveWindow) {
    effectiveWindow = Math.max(2048, suggested);
    downgraded = true;
    lastFailureEstimate = effectiveTokens;
   }
  },
 };
}

function normalize(value: number): number {
 if (!Number.isFinite(value) || value <= 0) return 200_000;
 return Math.floor(value);
}
