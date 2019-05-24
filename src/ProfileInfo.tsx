import React from 'react'
import { Shimmer } from 'office-ui-fabric-react';

interface Properties {
    authenticationHeaders?: Headers
}

export const ProfileInfo: React.FC<Properties> = ({ authenticationHeaders }: Properties) => {

    const [profile, setProfile] = React.useState();

    const getProfileAsync = async () => {
        if (authenticationHeaders) {
            var options = { method: "GET", headers: authenticationHeaders };

            const response = await fetch("https://graph.microsoft.com/v1.0/me", options);

            if (response.status === 200) {
                const me = await response.json();
                setProfile(me)
            }
        }
    }

    React.useEffect(() => { getProfileAsync() })

    if (profile) {
        return (
            <>
                <h1>Hello {profile.displayName} ({profile.userPrincipalName})</h1>
            </>)
    } else if (authenticationHeaders) {
        return (
            <Shimmer width="50%" />
        );
    } else {
        return <></>
    }
}