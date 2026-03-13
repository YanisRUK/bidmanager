import { type ReactNode } from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Text,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
    ...shorthands.padding(tokens.spacingVerticalXXXL, tokens.spacingHorizontalXXL),
    gap: tokens.spacingVerticalM,
  },
  icon: {
    fontSize: "48px",
    color: tokens.colorNeutralForeground4,
  },
});

interface EmptyStateProps {
  icon?: ReactNode;
  title: string;
  description?: string;
  action?: ReactNode;
}

export function EmptyState({ icon, title, description, action }: EmptyStateProps) {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      {icon && <div className={styles.icon}>{icon}</div>}
      <Text size={500} weight="semibold">
        {title}
      </Text>
      {description && (
        <Text size={300} style={{ color: tokens.colorNeutralForeground3, maxWidth: "400px" }}>
          {description}
        </Text>
      )}
      {action}
    </div>
  );
}
