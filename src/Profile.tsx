import React from 'react'
import { Shimmer, ShimmerElementType } from 'office-ui-fabric-react';

interface Properties {
    fetchWithScope: (scope: string, query: string) => Promise<Response | undefined>
}

export const Profile: React.FC<Properties> = ({ fetchWithScope }: Properties) => {

    const [profile, setProfile] = React.useState();

    if (profile) {
        return (
            <>
                <h1>{profile.displayName}</h1>
            </>)
    } else {
        const getProfileAsync = async () => {
            const graphQuery = "https://graph.microsoft.com/v1.0/me";
            const scope = "https://graph.microsoft.com/User.Read"

            const response = await fetchWithScope(scope, graphQuery);

            if (response) {
                const me = await response.json();
                setProfile(me)
            }
        }

        getProfileAsync()

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
    }
}