import { UserAgentApplication, AuthError, AuthResponse } from "msal";

declare global {
    interface Window { msalInstance: UserAgentApplication }
}

export const createUserAgentApplication = () => {

    if (!window.msalInstance) {
        const msalConfig = {
            auth: {
                clientId: 'c454a5cb-2667-4927-820a-89a1e25f0f8d',
                authority: "https://login.microsoftonline.com/common",
            }
        }

        window.msalInstance = new UserAgentApplication(msalConfig)

        const redirectCallback = (authErr: AuthError, response?: AuthResponse) => {
            console.log(authErr);
            console.log(response);
        }

        window.msalInstance.handleRedirectCallback(redirectCallback)
    }

    return window.msalInstance
}