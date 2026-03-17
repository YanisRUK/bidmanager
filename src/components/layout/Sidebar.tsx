import React from "react";
import { NavLink, useLocation } from "react-router-dom";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Tooltip,
  Divider,
} from "@fluentui/react-components";
import {
  GridRegular,
  AddSquareRegular,
  SettingsRegular,
  BuildingRegular,
} from "@fluentui/react-icons";
import { useAuth } from "../../context/AuthContext";

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

const useStyles = makeStyles({
  sidebar: {
    display: "flex",
    flexDirection: "column",
    width: "220px",
    minWidth: "220px",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    ...shorthands.overflow("hidden", "auto"),
    flexShrink: 0,
  },
  logoSection: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    ...shorthands.padding(tokens.spacingVerticalL, tokens.spacingHorizontalL),
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  logoIcon: {
    width: "32px",
    height: "32px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
    overflow: "hidden",
  },
  logoImg: {
    width: "32px",
    height: "32px",
    objectFit: "contain",
  },
  logoText: {
    display: "flex",
    flexDirection: "column",
    minWidth: 0,
  },
  nav: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalS),
    flexGrow: 1,
    gap: "2px",
  },
  navGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
  },
  navGroupLabel: {
    ...shorthands.padding(
      tokens.spacingVerticalS,
      tokens.spacingHorizontalM,
      tokens.spacingVerticalXS
    ),
    color: tokens.colorNeutralForeground3,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  navItem: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    textDecoration: "none",
    color: tokens.colorNeutralForeground2,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightRegular,
    lineHeight: tokens.lineHeightBase300,
    transition: "background-color 0.1s ease, color 0.1s ease",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      color: tokens.colorNeutralForeground1,
    },
  },
  navItemActive: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    fontWeight: tokens.fontWeightSemibold,
    ":hover": {
      backgroundColor: tokens.colorBrandBackground2Hover,
      color: tokens.colorBrandForeground1,
    },
  },
  navItemIcon: {
    fontSize: "18px",
    flexShrink: 0,
    color: "inherit",
  },
  navItemLabel: {
    flexGrow: 1,
    minWidth: 0,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  divider: {
    ...shorthands.margin(tokens.spacingVerticalS, 0),
  },
  userSection: {
    ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalL),
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  avatar: {
    width: "32px",
    height: "32px",
    borderRadius: "50%",
    backgroundColor: tokens.colorBrandBackground,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: tokens.colorNeutralForegroundOnBrand,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    flexShrink: 0,
  },
  userInfo: {
    minWidth: 0,
  },
});

// ---------------------------------------------------------------------------
// Nav item type
// ---------------------------------------------------------------------------

interface NavItemDef {
  label: string;
  to: string;
  icon: React.ReactNode;
  end?: boolean;
  /** Additional path prefixes that should also trigger the active state */
  alsoActiveFor?: string[];
}

const primaryNavItems: NavItemDef[] = [
  {
    label: "Dashboard",
    to: "/",
    icon: <GridRegular />,
    end: true,
  },
  {
    label: "New Bid",
    to: "/new-bid",
    icon: <AddSquareRegular />,
  },
];

const workspaceNavItems: NavItemDef[] = [
  {
    label: "Bid Workspaces",
    to: "/bid-register",
    icon: <BuildingRegular />,
    alsoActiveFor: ["/workspace"],
  },
];

const adminNavItems: NavItemDef[] = [
  {
    label: "Admin",
    to: "/admin",
    icon: <SettingsRegular />,
  },
];

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

function SidebarNavItem({ item }: { item: NavItemDef }) {
  const styles = useStyles();
  const location = useLocation();

  const isActive = item.end
    ? location.pathname === item.to
    : location.pathname.startsWith(item.to) ||
      (item.alsoActiveFor?.some((p) => location.pathname.startsWith(p)) ?? false);

  return (
    <Tooltip content={item.label} relationship="label" positioning="after">
      <NavLink
        to={item.to}
        end={item.end}
        className={`${styles.navItem} ${isActive ? styles.navItemActive : ""}`}
        aria-current={isActive ? "page" : undefined}
      >
        <span className={styles.navItemIcon}>{item.icon}</span>
        <span className={styles.navItemLabel}>{item.label}</span>
      </NavLink>
    </Tooltip>
  );
}

export function Sidebar() {
  const styles = useStyles();
  const { user } = useAuth();

  const initials = user?.fullName
    ? user.fullName
        .split(" ")
        .slice(0, 2)
        .map((n) => n[0])
        .join("")
        .toUpperCase()
    : "?";

  return (
    <aside className={styles.sidebar} aria-label="Main navigation">
      {/* Logo */}
      <div className={styles.logoSection}>
        <div className={styles.logoIcon}>
          <img src="/ricoh-logo-new.png" alt="Ricoh" className={styles.logoImg} />
        </div>
        <div className={styles.logoText}>
          <Text weight="semibold" size={300}>
            Ricoh
          </Text>
          <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
            Bid Manager
          </Text>
        </div>
      </div>

      {/* Navigation */}
      <nav className={styles.nav}>
        {/* Primary */}
        <div className={styles.navGroup}>
          <Text className={styles.navGroupLabel} size={100} weight="semibold">
            Main
          </Text>
          {primaryNavItems.map((item) => (
            <SidebarNavItem key={item.to} item={item} />
          ))}
        </div>

        <Divider className={styles.divider} />

        {/* Bid Workspaces */}
        <div className={styles.navGroup}>
          <Text className={styles.navGroupLabel} size={100} weight="semibold">
            Bid Workspaces
          </Text>
          {workspaceNavItems.map((item) => (
            <SidebarNavItem key={item.to} item={item} />
          ))}
        </div>

        <Divider className={styles.divider} />

        {/* Admin */}
        <div className={styles.navGroup}>
          {adminNavItems.map((item) => (
            <SidebarNavItem key={item.to} item={item} />
          ))}
        </div>
      </nav>

      {/* User section */}
      <div className={styles.userSection}>
        <div className={styles.avatar} aria-hidden="true">
          {initials}
        </div>
        <div className={styles.userInfo}>
          <Text
            size={200}
            weight="semibold"
            style={{ display: "block", color: tokens.colorNeutralForeground1 }}
          >
            {user?.fullName ?? "Unknown"}
          </Text>
          <Text
            size={100}
            style={{
              display: "block",
              color: tokens.colorNeutralForeground3,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {user?.email ?? ""}
          </Text>
        </div>
      </div>
    </aside>
  );
}
