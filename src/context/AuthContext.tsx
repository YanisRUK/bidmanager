/**
 * AuthContext
 *
 * Wraps the Power Apps SDK PowerProvider and exposes the current user and
 * authentication state throughout the app.
 *
 * In production this relies on @microsoft/powerapps-code-apps being available
 * inside the Power Apps runtime. During local development (npm run dev) we
 * fall back to a mock user so the UI is fully explorable without the SDK.
 */

import {
  createContext,
  useContext,
  useEffect,
  useState,
  type ReactNode,
} from "react";
import type { DataverseUser } from "../types/dataverse";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface AuthState {
  isAuthenticated: boolean;
  isLoading: boolean;
  user: DataverseUser | null;
  /** Error message if auth failed */
  error: string | null;
}

export interface AuthContextValue extends AuthState {
  /** Sign out (no-op in dev mode) */
  signOut: () => void;
}

// ---------------------------------------------------------------------------
// Context
// ---------------------------------------------------------------------------

const AuthContext = createContext<AuthContextValue | null>(null);

// ---------------------------------------------------------------------------
// Dev-mode mock user
// ---------------------------------------------------------------------------

const DEV_USER: DataverseUser = {
  id: "00000000-0000-0000-0000-000000000001",
  fullName: "Dev User",
  email: "dev@ricoh.co.uk",
  azureObjectId: "00000000-0000-0000-0000-000000000001",
};

// ---------------------------------------------------------------------------
// Provider
// ---------------------------------------------------------------------------

interface AuthProviderProps {
  children: ReactNode;
}

export function AuthProvider({ children }: AuthProviderProps) {
  const [state, setState] = useState<AuthState>({
    isAuthenticated: false,
    isLoading: true,
    user: null,
    error: null,
  });

  useEffect(() => {
    // Attempt to load the Power Apps SDK user context.
    // The SDK is only available inside the Power Apps runtime; we gracefully
    // degrade to the dev mock when it is absent.
    async function initAuth() {
      try {
        // Dynamic import so the bundle doesn't hard-fail in dev when the
        // package isn't installed. Cast to unknown to avoid type errors when
        // the package is absent.
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore – package only exists inside the Power Apps runtime
        const sdk = await (import("@microsoft/powerapps-code-apps") as Promise<unknown>).catch(
          () => null
        ) as Record<string, unknown> | null;

        if (sdk) {
          // SDK present — retrieve the authenticated user via the context API.
          // The exact API surface depends on the SDK version; adjust if needed.
          const getPowerAppsContext = sdk["getPowerAppsContext"] as (() => Promise<any>) | undefined;
          const context = await getPowerAppsContext?.();
          const sdkUser = context?.user;

          if (sdkUser) {
            setState({
              isAuthenticated: true,
              isLoading: false,
              user: {
                id: sdkUser.id ?? "",
                fullName: sdkUser.displayName ?? sdkUser.name ?? "",
                email: sdkUser.email ?? "",
                azureObjectId: sdkUser.azureObjectId,
              },
              error: null,
            });
            return;
          }
        }

        // SDK absent or user not available — use dev mock.
        console.info(
          "[AuthProvider] Power Apps SDK not available. Using dev mock user."
        );
        setState({
          isAuthenticated: true,
          isLoading: false,
          user: DEV_USER,
          error: null,
        });
      } catch (err) {
        const message =
          err instanceof Error ? err.message : "Authentication failed";
        setState({
          isAuthenticated: false,
          isLoading: false,
          user: null,
          error: message,
        });
      }
    }

    initAuth();
  }, []);

  function signOut() {
    setState({
      isAuthenticated: false,
      isLoading: false,
      user: null,
      error: null,
    });
  }

  return (
    <AuthContext.Provider value={{ ...state, signOut }}>
      {children}
    </AuthContext.Provider>
  );
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error("useAuth must be used within <AuthProvider>");
  }
  return ctx;
}
