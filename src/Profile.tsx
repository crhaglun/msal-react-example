import React from 'react'
import { Shimmer, ShimmerElementType } from 'office-ui-fabric-react';

interface Properties {
    scopedQuery: (scope: string, query: string) => Promise<Response | undefined> 
}

async function getProfileAsync(scopedQuery: (scope: string, query: string) => Promise<Response | undefined>, setProfile: React.Dispatch<any>) {
    const graphQuery = "https://graph.microsoft.com/v1.0/me";
    const scope = "https://graph.microsoft.com/User.Read"

    const response = await scopedQuery(scope, graphQuery);

    if (response)
    {
        const me = await response.json();
        setProfile(me)
    }
}

export const Profile: React.FC<Properties> = ( { scopedQuery } : Properties) => {

    const [profile, setProfile] = React.useState();

    if (profile) {
        return (
        <>
            <h1>Signed in as { profile.displayName }</h1>
        </>)
    } else {
        getProfileAsync(scopedQuery, setProfile)

        return (
            <Shimmer width="50%"
                shimmerElements={
                    [
                        { type: ShimmerElementType.circle },
                        { type: ShimmerElementType.gap, width: '2%' },
                        { type: ShimmerElementType.line }
                    ] }
            />
        );
    }
}