import {
  makeStyles,
  shorthands,
  tokens,
  Text,
  Button,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    ...shorthands.padding(tokens.spacingVerticalL),
  },
});

interface ErrorStateProps {
  message: string;
  onRetry?: () => void;
}

export function ErrorState({ message, onRetry }: ErrorStateProps) {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <MessageBar intent="error">
        <MessageBarBody>
          <Text weight="semibold">Something went wrong</Text>
          {" — "}
          <Text>{message}</Text>
          {onRetry && (
            <Button
              appearance="transparent"
              size="small"
              onClick={onRetry}
              style={{ marginLeft: "8px" }}
            >
              Retry
            </Button>
          )}
        </MessageBarBody>
      </MessageBar>
    </div>
  );
}
