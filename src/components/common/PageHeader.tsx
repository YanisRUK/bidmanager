import { type ReactNode } from "react";
import {
  makeStyles,
  tokens,
  Text,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    marginBottom: tokens.spacingVerticalXL,
    flexWrap: "wrap",
    gap: tokens.spacingVerticalM,
  },
  titleGroup: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
});

interface PageHeaderProps {
  title: string;
  subtitle?: string;
  actions?: ReactNode;
}

export function PageHeader({ title, subtitle, actions }: PageHeaderProps) {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.titleGroup}>
        <Text as="h1" size={700} weight="semibold">
          {title}
        </Text>
        {subtitle && (
          <Text size={300} style={{ color: tokens.colorNeutralForeground3 }}>
            {subtitle}
          </Text>
        )}
      </div>
      {actions && <div className={styles.actions}>{actions}</div>}
    </div>
  );
}
