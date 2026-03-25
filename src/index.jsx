import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider } from "@fluentui/react-components";
import App from "./App";
import { OfficeProvider } from "./context/OfficeContext";
import { AuthProvider } from "./context/AuthContext";

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

if (root) {
  root.render(
    <AuthProvider>
      <FluentProvider>
        <OfficeProvider>
          <App />
        </OfficeProvider>
      </FluentProvider>
    </AuthProvider>
  );
}