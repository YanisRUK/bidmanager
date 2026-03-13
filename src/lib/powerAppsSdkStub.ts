/**
 * Power Apps SDK stub for local development.
 *
 * This module is swapped in for @microsoft/powerapps-code-apps at dev time
 * via a Vite alias. It exports a no-op API that lets the app run locally
 * without the real SDK being installed.
 *
 * In production (Power Apps runtime) the real package is provided by the host.
 */

export async function getPowerAppsContext() {
  return null;
}
