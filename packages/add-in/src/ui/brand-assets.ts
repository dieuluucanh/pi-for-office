/**
 * Brand asset URLs for the taskpane UI.
 *
 * Imported through Vite (not the `public/` folder) so the emitted URL is
 * content-hashed and prefixed with the configured base path — `/` in dev and
 * `/pi-for-office/` on the GitHub Pages production build. The stable-named
 * copies in `public/assets/` remain the source for the Office manifests,
 * favicons, and the static landing pages.
 */

import icon32 from "../../assets/icon-32.png";
import icon80 from "../../assets/icon-80.png";

/** Small icon for inline marks (status bar, ~14px rendered). */
export const BRAND_ICON_SMALL: string = icon32;

/** Larger icon for hero/empty-state marks (~48px rendered). */
export const BRAND_ICON_MARK: string = icon80;
