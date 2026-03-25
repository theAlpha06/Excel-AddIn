import React from "react";
import { useAuth } from "../hooks/useAuth";
import { FontWeights, Persona, PersonaSize, Stack, mergeStyles } from "@fluentui/react";
import { PersonaPresence } from "@fluentui/react/lib/PersonaPresence";

const Profile = () => {
    const { user } = useAuth();
    
    const profileContainerStyles = mergeStyles({
        display: 'flex',
        justifyContent: 'center',
        padding: '20px',
    });

    return (
        <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }} className={profileContainerStyles}>
          <Persona
            imageUrl={user?.picture}
            text={user?.name || 'Salesforce User'}
            secondaryText={user?.email}
            size={PersonaSize.size72}
            presence={PersonaPresence.online}
            styles={{
              root: { margin: '0 auto' },
              primaryText: { 
                fontWeight: FontWeights.semibold,
                fontSize: '16px',
                color: '#323130'
              },
              secondaryText: { 
                color: '#605e5c',
                fontSize: '14px'
              },
              tertiaryText: {
                color: '#0078d4',
                fontSize: '12px',
                fontWeight: FontWeights.semibold
              }
            }}
          />
        </Stack>
    );
}

export default Profile;