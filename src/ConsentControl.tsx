import React from 'react'
import { AuthResponse } from 'msal'
import { Link, Label } from 'office-ui-fabric-react';

interface Properties {
    getAccessToken: (scope: string) => Promise<AuthResponse>,
    requestAccessToken: (scope: string) => void,
    description: string
    scope: string
}

export const ConsentControl: React.FC<Properties> = ({ getAccessToken, requestAccessToken, description, scope }: Properties) => {

    const [status, setStatus] = React.useState();

    const getScopeStatus = async () => {
        try {
            await getAccessToken(scope);
            setStatus("OK")
        }
        catch (err) {
            setStatus(err.errorCode)
        }
    }

    const getConsent = async () => {
        try {
            await requestAccessToken(scope)
        }
        catch (err) {
            setStatus(err.errorCode)
        }
    }

    if (status === "OK") {
        return (
            <>
                <Label>{description} : {status}</Label>
            </>
        )
    }
    else if (status) {
        return (
            <>
                <Label>{description} : {status}</Label>
                <Link onClick={() => getConsent()}>Get consent</Link>
            </>
        )
    } else {
        getScopeStatus()
        return (
            <Label>{description} : checking...</Label>
        )
    }
}