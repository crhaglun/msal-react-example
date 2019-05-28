import { AuthResponse, AuthError, UserAgentApplication } from "msal";
import React from "react";

interface AuthenticationState {
    readonly description: string
    readonly requestAccessToken: () => void
    readonly authResponse?: AuthResponse | AuthError
    readonly authHeaders?: Headers
}

export function useAuthenticationState(userAgentApplication: UserAgentApplication, description: string, scopes: string[]): AuthenticationState {

    const request = { scopes }

    const requestAccessToken = () => {
        try {
            // Use redirect to aquire a new access token; simplifies our app state
            // because we don't need to keep track of pending actions
            userAgentApplication.acquireTokenRedirect(request);
        } catch (err) {
            if (err.errorCode === "user_login_error") {
                // If we fail because the user is not signed in, redirect to the login
                // flow instead
                userAgentApplication.loginRedirect(request);
            } else {
                setAuthState((currentState: AuthenticationState) => { return { ...currentState, authResponse: err } })
            }
        }
    }

    const [authState, setAuthState] = React.useState<AuthenticationState>({ description, requestAccessToken });

    React.useEffect(() => {
        // Try to retrieve the currently cached access token and prepare a Headers object
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
    }, [authState, request, userAgentApplication])


    return authState;
}