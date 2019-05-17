import React from 'react'
import { Shimmer, ShimmerElementType } from 'office-ui-fabric-react/lib/Shimmer';

interface Properties {
    headers: Headers
}

async function getUnreadMailCountAsync(headers: Headers, setCount: React.Dispatch<any>) {

    var options = {
        method: "GET",
        headers: headers
    };

    // var graphEndpoint = "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages?$count=true&$select=id&$filter=isRead eq false";
    var query = "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages/$count?$filter=isRead eq false"

    const response = await fetch(query, options);

    if (response.status == 200) {
        const unreadMail = await response.json();
        setCount(unreadMail['@odata.count'])
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