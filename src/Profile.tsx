import React from 'react'
import { Shimmer, ShimmerElementType } from 'office-ui-fabric-react';

interface Properties {
    authenticationHeaders?: Headers
}

export const Profile: React.FC<Properties> = ({ authenticationHeaders }: Properties) => {

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
                <h1>Hello {profile.displayName}</h1>
            </>)
    } else if (authenticationHeaders) {
        return (
            <Shimmer width="50%"
                shimmerElements={
                    [
                        { type: ShimmerElementType.circle },
                        { type: ShimmerElementType.gap, width: '2%' },
                        { type: ShimmerElementType.line }
                    ]}
            />
        );
    } else {
        return <></>
    }
}