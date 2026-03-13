/**
 * AuthContext
 *
 * Wraps the official @microsoft/power-apps SDK and exposes the current user
 * and authentication state throughout the app.
 *
 * In production this runs inside the Power Apps host which provides the real
 * context. During local development (npm run dev via the powerApps() Vite
 * plugin) it falls back to a mock user so the UI is fully explorable.
 */

import {
  createContext,
  useContext,
  useEffect,
  useState,
  type ReactNode,
} from "react";
import * as app from "@microsoft/power-apps/app";
import type { DataverseUser } from "../types/dataverse";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface AuthState {
  isAuthenticated: boolean;
  isLoading: boolean;
  user: DataverseUser | null;
  error: string | null;
}

export interface AuthContextValue extends AuthState {
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
    async function initAuth() {
      try {
        const context = await app.getContext();
        const sdkUser = context?.user;

        if (sdkUser?.userPrincipalName) {
          setState({
            isAuthenticated: true,
            isLoading: false,
            user: {
              id: sdkUser.objectId ?? "",
              fullName: sdkUser.fullName ?? sdkUser.userPrincipalName,
              email: sdkUser.userPrincipalName,
              azureObjectId: sdkUser.objectId,
            },
            error: null,
          });
        } else {
          // Running outside Power Apps host (bare browser) — use dev mock
          console.info(
            "[AuthProvider] No Power Apps user context. Using dev mock user."
          );
          setState({
            isAuthenticated: true,
            isLoading: false,
            user: DEV_USER,
            error: null,
          });
        }
      } catch {
        // getContext() throws when not inside the Power Apps host — use dev mock
        console.info(
          "[AuthProvider] Power Apps host not detected. Using dev mock user."
        );
        setState({
          isAuthenticated: true,
          isLoading: false,
          user: DEV_USER,
          error: null,
        });
      }
    }

    initAuth();
  }, []);

  function signOut() {
    setState({ isAuthenticated: false, isLoading: false, user: null, error: null });
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
  if (!ctx) throw new Error("useAuth must be used within <AuthProvider>");
  return ctx;
}
