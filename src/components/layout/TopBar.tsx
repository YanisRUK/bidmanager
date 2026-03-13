
import { useLocation } from "react-router-dom";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Button,
  Tooltip,
} from "@fluentui/react-components";
import {
  AlertRegular,
  QuestionCircleRegular,
} from "@fluentui/react-icons";

// ---------------------------------------------------------------------------
// Route → page title mapping
// ---------------------------------------------------------------------------

const PAGE_TITLES: Record<string, string> = {
  "/":                     "Dashboard",
  "/bid-register":         "Bid Register",
  "/new-bid":              "New Bid",
  "/workspace":            "Bid Workspace",
  "/workspace/documents":  "Bid Documents",
  "/admin":                "Admin",
};

function getTitle(pathname: string): string {
  // Exact match first
  if (PAGE_TITLES[pathname]) return PAGE_TITLES[pathname];
  // Prefix match (e.g. /workspace/ws-001)
  const key = Object.keys(PAGE_TITLES)
    .filter((k) => k !== "/")
    .find((k) => pathname.startsWith(k));
  return key ? PAGE_TITLES[key] : "Bid Manager";
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  topBar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    height: "52px",
    ...shorthands.padding(0, tokens.spacingHorizontalXL),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    flexShrink: 0,
  },
  actions: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
});

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function TopBar() {
  const styles = useStyles();
  const location = useLocation();
  const title = getTitle(location.pathname);

  return (
    <header className={styles.topBar} role="banner">
      <Text size={500} weight="semibold">
        {title}
      </Text>

      <div className={styles.actions}>
        <Tooltip content="Notifications" relationship="label">
          <Button
            appearance="subtle"
            icon={<AlertRegular />}
            aria-label="Notifications"
          />
        </Tooltip>
        <Tooltip content="Help" relationship="label">
          <Button
            appearance="subtle"
            icon={<QuestionCircleRegular />}
            aria-label="Help"
          />
        </Tooltip>
      </div>
    </header>
  );
}
