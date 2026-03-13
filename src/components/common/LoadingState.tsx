import {
  makeStyles,
  shorthands,
  tokens,
  Spinner,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    ...shorthands.padding(tokens.spacingVerticalXXXL),
    gap: tokens.spacingVerticalM,
  },
});

interface LoadingStateProps {
  label?: string;
}

export function LoadingState({ label = "Loading..." }: LoadingStateProps) {
  const styles = useStyles();
  return (
    <div className={styles.root} role="status" aria-live="polite">
      <Spinner size="medium" label={label} />
    </div>
  );
}
