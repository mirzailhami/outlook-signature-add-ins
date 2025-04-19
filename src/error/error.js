import * as React from "react";
import { createRoot } from "react-dom/client";
import { MessageBar, MessageBarTitle, MessageBarBody, MessageBarActions, Button, FluentProvider, webLightTheme, makeStyles } from "@fluentui/react-components";
import { DismissRegular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  container: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center"
  },
  messageBar: {
    width: "100%",
  },
});

const ErrorDialog = ({ message }) => {
  const styles = useStyles();
  const handleClose = () => {
    try {
      Office.context.ui.messageParent("close");
    } catch (error) {
      console.error("Failed to close dialog:", error);
    }
  };

  return (
    <div className={styles.container}>
      <MessageBar intent="error" role="alertdialog" className={styles.messageBar}>
        <MessageBarBody>
          {message || "An unexpected error occurred."}
        </MessageBarBody>
        <MessageBarActions>
          <Button appearance="filled" onClick={handleClose} aria-label="Close error dialog">
            Close
          </Button>
        </MessageBarActions>
      </MessageBar>
    </div>
  );
};

Office.onReady(() => {
  const params = new URLSearchParams(window.location.search);
  const errorMessage = params.get("message") || "An error occurred.";
  const root = createRoot(document.getElementById("container"));
  root.render(
    <FluentProvider theme={webLightTheme}>
      <ErrorDialog message={errorMessage} />
    </FluentProvider>
  );
});