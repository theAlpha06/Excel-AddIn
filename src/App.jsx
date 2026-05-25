import React, { useState, useEffect } from "react";
import { mergeStyles } from "@fluentui/react";
import Login from "./pages/Login";
import Home from "./pages/Home";
import { useOfficeContext } from "./hooks/useOfficeContext";
import { useAuth } from "./hooks/useAuth";
import { useLoader } from "./hooks/useLoader";
import Loader from "./components/Loader";
import { handleSheetChange, registerPayloadSetter } from "./utils/diff";
import { useDiff } from "./context/DiffContext";

export let sheetChangeHandler = null;

const App = () => {
  const [error, setError] = useState(null);
  const [orgType, setOrgType] = useState(null);
  const { isOfficeReady, officeError } = useOfficeContext();
  const { setUser, isConnected, setIsConnected } = useAuth();
  const { isLoading, setIsLoading } = useLoader();
  const { setPendingPayload } = useDiff();

  const errorMessageStyles = mergeStyles({
    padding: "20px",
    color: "red",
  });

  const loadingMessageStyles = mergeStyles({
    padding: "20px",
  });

  useEffect(() => {
    if (isOfficeReady) {
      checkConnectionStatus();
      registerChangeListener();
    }
  }, [isOfficeReady]);

  useEffect(() => {
    registerPayloadSetter(setPendingPayload);
    console.log("✅ Registered payload setter in App");
  }, [setPendingPayload]);

  async function registerChangeListener() {
    await Excel.run(async (context) => {
      context.workbook.worksheets.onChanged.add(handleSheetChange);
      await context.sync();
    });
  }

  const checkConnectionStatus = async () => {
    try {
      if (!isOfficeReady || typeof Office === "undefined" || !Office.context) {
        setError("Office context is not ready");
        setIsLoading(false);
        return;
      }

      setIsLoading(true);
      await Office.context.document.settings.refreshAsync();
      const accessToken = Office.context.document.settings.get("salesforce_access_token");
      const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");
      if (accessToken && instanceUrl) {
        try {
          const response = await fetch("http://localhost:5000/salesforce/check-status", {
            method: "GET",
            headers: {
              "sf-instance-url": instanceUrl,
              "sf-access-token": accessToken,
              "Content-Type": "application/json"
            },
          });
          const data = await response.json();
          if (response.ok && data.status === "connected") {
            setIsConnected(true);
            if (data.versions) {
              setUser(data.versions);
            }
          } else {
            // handleLogout();
            // TO BE HANDLED LATER
          }
        } catch (apiError) {
          setError("Could not verify salesforce connection. Please check your internet connection");
        }
      }
    } catch (error) {
      setError("Failed to check connection status");
    } finally {
      setIsLoading(false);
    }
  };

  if (officeError) {
    return <div className={errorMessageStyles}>Error: {officeError}</div>;
  }

  if (!isOfficeReady) {
    return <div className={loadingMessageStyles}>Loading Office context...</div>;
  }

  return (
    <>
      {isLoading ? (
        <Loader />
      ) : (
        <div>{isConnected ? <Home /> : <Login orgType={orgType} setOrgType={setOrgType} />}</div>
      )}
    </>
  );
};

export default App;