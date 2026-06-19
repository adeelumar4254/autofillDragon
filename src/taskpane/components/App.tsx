import * as React from "react";
import { makeStyles } from "@fluentui/react-components";
import AppRouter from "../Router/AppRouter";
import { SnackbarProvider } from "../context/SnackbarContext";
const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});


const App: React.FC = () => {
  const styles = useStyles();


  return (
    <div className={styles.root}>
     
      <SnackbarProvider>
        <AppRouter />
        {/* <ErrorReportTest /> */}
      </SnackbarProvider>

    </div >
  );
};

export default App;
