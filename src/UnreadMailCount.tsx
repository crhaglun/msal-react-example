import React from 'react'
import { Shimmer } from 'office-ui-fabric-react';

interface Properties {
    headers: Headers
}

async function getUnreadMailCountAsync(headers: Headers, setCount: React.Dispatch<any>) {

    var options = {
        method: "GET",
        headers: headers
    };

    var graphQuery = "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages/$count?$filter=isRead eq false"

    const response = await fetch(graphQuery, options);

    if (response.status === 200) {
        const unreadMail = await response.json();
        setCount(unreadMail)
    }
    else {
        setCount(null);
    }
}

export const UnreadMailCount: React.FC<Properties> = ({ headers }: Properties) => {

    const [count, setCount] = React.useState();

    if (headers && count === undefined) {
        getUnreadMailCountAsync(headers, setCount)

        return (
            <Shimmer width="50%" />
        );
    }

    if (count > 0) {
        return (<h2>You have { count } unread mail :-(</h2>)
    } else if (count === 0) {
        return (<h2>You have no unread mail! :-)</h2>)
    } else {
        return (<h2>Could not check mail status :-/</h2>)
    }
}