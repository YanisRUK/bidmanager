import { Outlet } from "react-router-dom";
import { makeStyles, shorthands, tokens } from "@fluentui/react-components";
import { Sidebar } from "./Sidebar";
import { TopBar } from "./TopBar";

const useStyles = makeStyles({
  root: {
    display: "flex",
    height: "100vh",
    width: "100vw",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  main: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minWidth: 0,
    overflow: "hidden",
  },
  content: {
    flexGrow: 1,
    ...shorthands.overflow("hidden", "auto"),
    ...shorthands.padding(tokens.spacingVerticalL, tokens.spacingHorizontalXL),
  },
});

export function AppLayout() {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Sidebar />
      <div className={styles.main}>
        <TopBar />
        <main className={styles.content} id="main-content" tabIndex={-1}>
          <Outlet />
        </main>
      </div>
    </div>
  );
}
