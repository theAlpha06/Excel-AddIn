import React, { useState } from "react";
import { ChoiceGroup, DefaultButton, mergeStyles } from "@fluentui/react";
import { useOfficeContext } from "../hooks/useOfficeContext";

const Login = ({orgType, setOrgType}) => {
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState(null);
    const [dialog, setDialog] = useState(null);
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

    const processMessage = (message) => {
        if (dialog) {
            dialog.close();
        }
        const { access_token, instance_url } = JSON.parse(message.message);
        if (access_token && instance_url) {
            try {
                Office.context.document.settings.set("salesforce_access_token", access_token);
                Office.context.document.settings.set("salesforce_instance_url", instance_url);
                Office.context.document.settings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        console.error("Failed to save settings:", result.error.message);
                        setError(`Failed to save connection settings: ${result.error.message}`);
                    } else {
                        setError(null);
                        setIsLoading(false);
                        window.location.href = "taskpane.html";
                    }
                });
            } catch (error) {
                setError(`Error processing authentication response: ${error.message}`);
            }
        } else {
            setError("Authentication failed: No access token received");
        }
    };

    const processDialogEvent = (message) => {
        setIsLoading(false);
        setError("Dialog Box was closed by user");
        if (dialog) {
            dialog.close();
        }
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

            const authUrl =
                `${LOGIN_URL}/services/oauth2/authorize?` +
                `response_type=token` +
                `&client_id=${REACT_APP_CONSUMER_KEY}` +
                `&redirect_uri=${encodeURIComponent(REACT_APP_REDIRECT_URI)}`;

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
                        setDialog(dialogInstance);

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