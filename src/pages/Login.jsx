import React from "react";
import { ChoiceGroup, DefaultButton } from "@fluentui/react";

const Login = ({orgType, setOrgType}) => {

    const options = [
        { key: 'sandbox', text: 'Sandbox' },
        { key: 'production', text: 'Production' }
    ];

    const handleOrgTypeChange = (ev, option) => {
        if (option) {
            setOrgType(option.key);
        }
    };
    
    return (
        <div  style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100vh' }}>
            <div style={{ marginBottom: '20px' }}>
                <ChoiceGroup
                    selectedKey={orgType}
                    options={options}
                    onChange={handleOrgTypeChange}
                    label="Select Organization"
                />
            </div>
            <div>
                <DefaultButton
                text="Connect to Salesforce"
                iconProps={{ iconName: "PlugConnected" }}
                primary
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