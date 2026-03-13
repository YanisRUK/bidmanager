/**
 * App.tsx — Root component
 *
 * Provides:
 *  - Fluent UI v9 theme (FluentProvider)
 *  - AuthProvider (Power Apps SDK / dev mock)
 *  - React Router with the AppLayout shell and all page routes
 */

import React from "react";
import {
  FluentProvider,
  webLightTheme,
  Spinner,
  makeStyles,
  tokens,
  Text,
} from "@fluentui/react-components";
import {
  BrowserRouter,
  Routes,
  Route,
  Navigate,
} from "react-router-dom";

import { AuthProvider, useAuth } from "./context/AuthContext";
import { AppLayout } from "./components/layout/AppLayout";
import { DashboardPage } from "./pages/Dashboard/DashboardPage";
import { BidRegisterPage } from "./pages/BidRegister/BidRegisterPage";
import { NewBidPage } from "./pages/NewBid/NewBidPage";
import { BidWorkspacePage } from "./pages/BidWorkspace/BidWorkspacePage";
import { AdminPage } from "./pages/Admin/AdminPage";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  loadingScreen: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100vh",
    gap: tokens.spacingVerticalL,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  errorScreen: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100vh",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalXL,
    textAlign: "center",
    backgroundColor: tokens.colorNeutralBackground1,
  },
});

// ---------------------------------------------------------------------------
// Auth gate — shows loading / error before rendering routes
// ---------------------------------------------------------------------------

function AuthGate({ children }: { children: React.ReactNode }) {
  const styles = useStyles();
  const { isLoading, isAuthenticated, error } = useAuth();

  if (isLoading) {
    return (
      <div className={styles.loadingScreen} role="status" aria-live="polite">
        <Spinner size="large" label="Initialising Ricoh Bid Manager..." />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.errorScreen} role="alert">
        <Text size={700} weight="bold">Authentication Error</Text>
        <Text size={400} style={{ color: tokens.colorNeutralForeground3 }}>
          {error}
        </Text>
        <Text size={300} style={{ color: tokens.colorNeutralForeground4 }}>
          Ensure this app is running inside Power Apps or that your dev
          environment is configured correctly.
        </Text>
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className={styles.errorScreen}>
        <Text size={700} weight="bold">Not Authenticated</Text>
        <Text size={400} style={{ color: tokens.colorNeutralForeground3 }}>
          Please open this app through Power Apps.
        </Text>
      </div>
    );
  }

  return <>{children}</>;
}

// ---------------------------------------------------------------------------
// Router
// ---------------------------------------------------------------------------

function AppRoutes() {
  return (
    <Routes>
      <Route element={<AppLayout />}>
        {/* Dashboard */}
        <Route index element={<DashboardPage />} />

        {/* Bid Register */}
        <Route path="bid-register" element={<BidRegisterPage />} />

        {/* New Bid intake form */}
        <Route path="new-bid" element={<NewBidPage />} />

        {/* Bid Workspace */}
        <Route path="workspace" element={<BidWorkspacePage />} />
        <Route path="workspace/:workspaceId" element={<BidWorkspacePage />} />

        {/* Admin */}
        <Route path="admin" element={<AdminPage />} />

        {/* Catch-all */}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Route>
    </Routes>
  );
}

// ---------------------------------------------------------------------------
// Root
// ---------------------------------------------------------------------------

export default function App() {
  return (
    <FluentProvider theme={webLightTheme}>
      <AuthProvider>
        <BrowserRouter>
          <AuthGate>
            <AppRoutes />
          </AuthGate>
        </BrowserRouter>
      </AuthProvider>
    </FluentProvider>
  );
}
