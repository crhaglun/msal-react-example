import React from 'react'
import { AuthResponse, AuthError } from 'msal'
import { Link, Label } from 'office-ui-fabric-react';

interface Properties {
    accessToken?: AuthResponse | AuthError,
    requestAccessToken: () => void,
    description: string
}

export const ConsentControl: React.FC<Properties> = ({ accessToken, requestAccessToken, description }: Properties) => {

    var statusText: string;
    var linkText: string | undefined;

    if (!accessToken) {
        statusText = "checking..."
    }
    else if (accessToken instanceof AuthError) {
        statusText = accessToken.errorCode

        switch (accessToken.errorCode) {
            case "token_renewal_error":
            case "login_required":
                linkText = "Sign in"
                break;

            case "consent_required":
                linkText = "Get consent"
                break;

            default:
                linkText = "Force sign-in"
        }
    }
    else {
        statusText = "OK"
    }

    if (linkText) {
        return (
            <>
                <Label>{description} : {statusText}</Label>
                <Link onClick={() => requestAccessToken()}>{linkText}</Link>
            </>
        )
    }
    else {
        return (
            <>
                <Label>{description} : {statusText}</Label>
            </>
        )
    }
}