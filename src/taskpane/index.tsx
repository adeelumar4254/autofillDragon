/* global document, Office, module, require */
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { HashRouter } from "react-router-dom";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {


  root?.render(
    <HashRouter>
      <App />
    </HashRouter>
  );
});

// For Hot Module Replacement
if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(
      <HashRouter>
        <NextApp />
      </HashRouter>
    );
  });
}