import React, { useState, useRef } from "react";
import { ChoiceGroup, DefaultButton, mergeStyles } from "@fluentui/react";
import { useOfficeContext } from "../hooks/useOfficeContext";

const Login = ({orgType, setOrgType}) => {
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState(null);
    const dialogRef = useRef(null);
    const authHandledRef = useRef(false);
    const { isOfficeReady } = useOfficeContext();

    const loginContainerStyles = mergeStyles({
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100vh',
        padding: '20px',
    });

    const choiceGroupWrapperStyles = mergeStyles({
        marginBottom: '20px',
        minWidth: '300px',
    });

    const errorDivStyles = mergeStyles({
        marginBottom: '20px',
        color: 'red',
        fontSize: '14px',
        fontWeight: '500',
    });

    const options = [
        { key: 'sandbox', text: 'Sandbox' },
        { key: 'production', text: 'Production' }
    ];

    const LOGIN_URL =
        orgType === "production" ? "https://login.salesforce.com" : "https://test.salesforce.com";

    const handleOrgTypeChange = (ev, option) => {
        if (option) {
            setOrgType(option.key);
        }
    };

    const REACT_APP_CONSUMER_KEY="3MVG9GCMQoQ6rpzQm797ZYHBww6Iye8XtvMqgddHtVjTr83ocGxvaeyGNJuUUf_dDrt688z_OV9wi8c2RC90w";
    const REACT_APP_REDIRECT_URI="https://localhost:3000/auth-callback.html";

    // Closing the auth dialog is a race nobody owns: auth-callback.html closes itself shortly
    // after messageParent, the user can dismiss it, and the Office host tears it down on its
    // own. close() on a dialog that is already gone throws
    // "Function close call failed, error code: 12006", so every close goes through here — at
    // most once, and never fatally. Letting it throw used to abort processMessage before the
    // access token was saved (and the catch block's own close() threw a second time), which is
    // why a successful first sign-in still landed back on the login screen.
    const closeDialog = () => {
        const dialog = dialogRef.current;
        if (!dialog) return;
        dialogRef.current = null;
        try {
            dialog.close();
        } catch (err) {
            console.warn("Auth dialog was already closed:", err?.message ?? err);
        }
    };

    const processMessage = async (message) => {
        authHandledRef.current = true;

        const { code, error: authError } = JSON.parse(message.message);

        if (authError) {
            setIsLoading(false);
            setError(`Authentication failed: ${authError}`);
            closeDialog();
            return;
        }

        if (!code) {
            setIsLoading(false);
            setError("Authentication failed: No authorization code received");
            closeDialog();
            return;
        }

        try {
            const tokenRes = await fetch("http://localhost:5000/salesforce/oauth/token", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    code,
                    redirect_uri: REACT_APP_REDIRECT_URI,
                    login_url: LOGIN_URL,
                }),
            });

            const data = await tokenRes.json();

            if (!tokenRes.ok || !data.access_token || !data.instance_url) {
                throw new Error(data.message || "Failed to exchange authorization code");
            }

            // Hold the token before touching the dialog. set() is in-memory and cannot fail on
            // a dialog teardown, so ordering it first means no close-time surprise can cost the
            // user the sign-in they just completed.
            Office.context.document.settings.set("salesforce_access_token", data.access_token);
            Office.context.document.settings.set("salesforce_instance_url", data.instance_url);

            closeDialog();

            Office.context.document.settings.saveAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Failed to save settings:", result.error.message);
                    setError(`Failed to save connection settings: ${result.error.message}`);
                    setIsLoading(false);
                } else {
                    setError(null);
                    setIsLoading(false);
                    window.location.href = "taskpane.html";
                }
            });
        } catch (error) {
            closeDialog();
            setIsLoading(false);
            setError(`Error processing authentication response: ${error.message}`);
        }
    };

    const processDialogEvent = () => {
        // The dialog is gone by the time this fires, so drop the handle: keeping it would let a
        // later close() call reach a dead dialog and throw 12006.
        dialogRef.current = null;

        if (authHandledRef.current) {
            // Expected: the dialog closes itself right after we already processed
            // its message (either via our own close() call above, or because the
            // Office dialog API auto-closes after messageParent). Nothing to report.
            return;
        }
        setIsLoading(false);
        setError("Sign-in was cancelled.");
    };

    const startOAuth = async () => {
        if (!isOfficeReady) {
            setError("Office context is not ready. Please wait for the add-in to initialize.");
            return;
        }

        if (typeof Office === 'undefined' || !Office.context || !Office.context.ui) {
            setError("Office API is not available.");
            return;
        }

        try {
            setIsLoading(true);
            setError(null);
            authHandledRef.current = false;

            const authUrl =
                `${LOGIN_URL}/services/oauth2/authorize?` +
                `response_type=code` +
                `&client_id=${REACT_APP_CONSUMER_KEY}` +
                `&redirect_uri=${encodeURIComponent(REACT_APP_REDIRECT_URI)}` +
                `&prompt=login`;

            Office.context.ui.displayDialogAsync(
                authUrl,
                { height: 80, width: 50, promptBeforeOpen: true },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to open dialog:", result.error.message);
                        setError(`Failed to open login dialog: ${result.error.message}`);
                        setIsLoading(false);
                    } else {
                        const dialogInstance = result.value;
                        dialogRef.current = dialogInstance;

                        dialogInstance.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                        dialogInstance.addEventHandler(
                            Office.EventType.DialogEventReceived,
                            processDialogEvent
                        );
                    }
                }
            );
        } catch (error) {
            setError(`Authorization failed: ${error.message}`);
            setIsLoading(false);
        }
    };
    
    return (
        <div className={loginContainerStyles}>
            <div className={choiceGroupWrapperStyles}>
                <ChoiceGroup
                    selectedKey={orgType}
                    options={options}
                    onChange={handleOrgTypeChange}
                    label="Select Organization"
                />
            </div>
            {error && <div className={errorDivStyles}>{error}</div>}
            <div>
                <DefaultButton
                    text="Connect to Salesforce"
                    iconProps={{ iconName: "PlugConnected" }}
                    onClick={startOAuth}
                    primary
                    disabled={!isOfficeReady || isLoading}
                    styles={{
                        root: {
                        minWidth: '180px',
                        height: '40px',
                        fontSize: '14px'
                        }
                    }}
                />
            </div>
        </div>
    );
};

export default Login;