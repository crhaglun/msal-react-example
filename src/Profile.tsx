import React from 'react'
import { Shimmer, ShimmerElementType } from 'office-ui-fabric-react';

interface Properties {
    headers: Headers
}

async function getProfileAsync(headers: Headers, setProfile: React.Dispatch<any>) {

    var options = {
        method: "GET",
        headers: headers
    };

    var graphEndpoint = "https://graph.microsoft.com/v1.0/me";

    const response = await fetch(graphEndpoint, options);
    const me = await response.json();

    setProfile(me)
}

export const Profile: React.FC<Properties> = ({ headers }: Properties) => {

    const [profile, setProfile] = React.useState();

    if (headers) {
        getProfileAsync(headers, setProfile)
    }

    if (profile) {
        return (<h1>Signed in as { profile.displayName }</h1>)
    } else {
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