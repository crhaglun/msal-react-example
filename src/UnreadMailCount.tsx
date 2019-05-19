import React from 'react'
import { Shimmer } from 'office-ui-fabric-react';

interface Properties {
    scopedQuery: (scope: string, query: string) => Promise<Response | undefined> 
}

async function getUnreadMailCountAsync(scopedQuery: (scope: string, query: string) => Promise<Response | undefined>, setCount: React.Dispatch<any>) {
    const graphQuery = "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages/$count?$filter=isRead eq false";
    const scope = "https://graph.microsoft.com/Mail.Read"

    const response = await scopedQuery(scope, graphQuery);

    if (response && response.status === 200) {
        const unreadMail = await response.json();
        setCount(unreadMail)
    }
    else {
        setCount(null);
    }
}

export const UnreadMailCount: React.FC<Properties> = ({ scopedQuery }: Properties) => {

    const [count, setCount] = React.useState();

    if (count === undefined) {
        getUnreadMailCountAsync(scopedQuery, setCount)

        return (
            <Shimmer width="50%" />
        );
    }
    else if (count > 0) {
        return (<h2>You have { count } unread mail :-(</h2>)
    } else if (count === 0) {
        return (<h2>You have no unread mail! :-)</h2>)
    } else {
        return (<h2>Could not check mail status :-/</h2>)
    }
}