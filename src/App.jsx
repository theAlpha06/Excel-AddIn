import React, { useState, useEffect } from "react";
import { mergeStyles } from "@fluentui/react";
import Login from "./pages/Login";
import Home from "./pages/Home";
import { useOfficeContext } from "./hooks/useOfficeContext";
import { useAuth } from "./hooks/useAuth";

const App = () => {
    const [isConnected, setIsConnected] = useState(false);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);
    const [orgType, setOrgType] = useState(null);
    const { isOfficeReady, officeError } = useOfficeContext();
    const { setUser } = useAuth();

    const errorMessageStyles = mergeStyles({
        padding: '20px',
        color: 'red',
    });

    const loadingMessageStyles = mergeStyles({
        padding: '20px',
    });

    useEffect(() => {
        if (isOfficeReady) {
            checkConnectionStatus();
        }
    }, [isOfficeReady]);

    const checkConnectionStatus = async () => {
        try {
            if (!isOfficeReady || typeof Office === 'undefined' || !Office.context) {
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
                            "Content-Type": "application/json",
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
                    }
                } catch (apiError) {
                    console.error("Error verifying token:", apiError);
                    setError("Could not verify Salesforce connection. Check your internet connection.");
                }
            }
        } catch (error) {
            console.error("Error checking connection status:", error);
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
        <div>
            {isConnected ? <Home /> : <Login orgType={orgType} setOrgType={setOrgType} />}
        </div>
    );
}

export default App;