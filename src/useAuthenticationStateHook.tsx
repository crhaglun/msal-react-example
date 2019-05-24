import { AuthResponse, AuthError, UserAgentApplication } from "msal";
import React from "react";

interface AuthenticationState {
    readonly description: string
    readonly authResponse?: AuthResponse | AuthError
    readonly authHeaders?: Headers
    readonly requestAccessToken: () => void
}

export function useAuthenticationState(userAgentApplication: UserAgentApplication, description: string, scopes: string[]): AuthenticationState {

    const request = {
        scopes: scopes
    }

    const requestAccessToken = () => {
        try {
            userAgentApplication.acquireTokenRedirect(request);
        } catch (err) {
            if (err.errorCode === "user_login_error") {
                userAgentApplication.loginRedirect(request);
            } else {
                setAuthState((currentState: AuthenticationState) => { return { ...currentState, authResponse: err } })
            }
        }
    }

    const [authState, setAuthState] = React.useState<AuthenticationState>(
        { description, requestAccessToken });

    React.useEffect(() => {
        if (!authState.authResponse) {
            const initializeAuthResponse = async () => {
                try {
                    const authResponse = await userAgentApplication.acquireTokenSilent(request)
                    const token = authResponse.accessToken;
                    const authHeaders = new Headers({ "Authorization": "Bearer " + token })

                    setAuthState((currentState: AuthenticationState) => { return { ...currentState, authResponse, authHeaders } })
                }
                catch (err) {
                    if (err instanceof AuthError) {
                        setAuthState((currentState: AuthenticationState) => { return { ...currentState, authResponse: err } })
                    }
                    else {
                        console.log(err)
                    }
                }
            }

            initializeAuthResponse()
        }
    }, [authState])


    return authState;
}