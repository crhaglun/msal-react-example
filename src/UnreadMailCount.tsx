import React from 'react'
import { Shimmer } from 'office-ui-fabric-react';

interface Properties {
    fetchWithScope: (scope: string, query: string) => Promise<Response | undefined>
}

export const UnreadMailCount: React.FC<Properties> = ({ fetchWithScope }: Properties) => {

    const [count, setCount] = React.useState();

    if (count === undefined) {
        const getUnreadMailCountAsync = async () => {
            const graphQuery = "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages/$count?$filter=isRead eq false";
            const scope = "https://graph.microsoft.com/Mail.Read"

            const response = await fetchWithScope(scope, graphQuery);

            if (response && response.status === 200) {
                const unreadMail = await response.json();
                setCount(unreadMail)
            }
            else {
                setCount(null);
            }
        }

        getUnreadMailCountAsync()

        return (
            <Shimmer width="50%" />
        );
    }
    else if (count > 0) {
        return (<h2>You have {count} unread mail :-(</h2>)
    } else if (count === 0) {
        return (<h2>You have no unread mail! :-)</h2>)
    } else {
        return (<h2>Could not check mail status :-/</h2>)
    }
}