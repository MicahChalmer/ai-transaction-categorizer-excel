import * as React from "react";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (_props: AppProps) => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <TextInsertion insertText={insertText} />
    </div>
  );
};

export default App;
