import React from "react";
import { useAuth } from "../hooks/useAuth";
import {
  DefaultButton,
  FontWeights,
  Persona,
  PersonaSize,
  Stack,
  mergeStyles,
} from "@fluentui/react";
import { PersonaPresence } from "@fluentui/react/lib/PersonaPresence";

const Profile = () => {
  const { user, logout, setIsConnected, setUser } = useAuth();

  const profileContainerStyles = mergeStyles({
    display: "flex",
    justifyContent: "center",
    padding: "20px",
  });

  const buttonStackTokens = {
    childrenGap: 12,
  };

  const handleLogout = () => {
    try {
      Office.context.document.settings.remove("salesforce_access_token");
      Office.context.document.settings.remove("salesforce_refresh_token");
      Office.context.document.settings.remove("salesforce_instance_url");

      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setIsConnected(false);
          setUser(null);
        } else {
          setError(`Failed to logout: ${result.error.message}`);
        }
      });
      logout();
    } catch (err) {
      setError(`Logout failed: ${err.message}`);
    }
  };

  return (
    <Stack>
      <Stack
        horizontalAlign="center"
        tokens={{ childrenGap: 12 }}
        className={profileContainerStyles}
      >
        <Persona
          imageUrl={user?.picture}
          text={user?.name || "Salesforce User"}
          secondaryText={user?.email}
          size={PersonaSize.size72}
          presence={PersonaPresence.online}
          styles={{
            root: { margin: "0 auto" },
            primaryText: {
              fontWeight: FontWeights.semibold,
              fontSize: "16px",
              color: "#323130",
            },
            secondaryText: {
              color: "#605e5c",
              fontSize: "14px",
            },
            tertiaryText: {
              color: "#0078d4",
              fontSize: "12px",
              fontWeight: FontWeights.semibold,
            },
          }}
        />
      </Stack>
      <Stack horizontal tokens={buttonStackTokens} horizontalAlign="center">
        <DefaultButton
          text="Disconnect"
          onClick={handleLogout}
          iconProps={{ iconName: "PlugDisconnected" }}
          styles={{
            root: {
              minWidth: "120px",
              borderColor: "#d83b01",
              color: "#d83b01",
            },
            rootHovered: {
              borderColor: "#a4262c",
              color: "#a4262c",
              backgroundColor: "#fdf6f6",
            },
          }}
        />
      </Stack>
    </Stack>
  );
};

export default Profile;
