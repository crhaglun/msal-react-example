import React from 'react'
import { Shimmer } from 'office-ui-fabric-react';

interface Properties {
    authenticationHeaders?: Headers
}

export const UnreadMailCount: React.FC<Properties> = ({ authenticationHeaders }: Properties) => {

    const [count, setCount] = React.useState();

    const getUnreadMailCountAsync = async () => {
        if (authenticationHeaders) {
            const unreadMailQuery = "https://graph.microsoft.com/v1.0/me/mailFolders('Inbox')/messages/$count?$filter=isRead eq false";

            var options = { method: "GET", headers: authenticationHeaders };

            const response = await fetch(unreadMailQuery, options);

            if (response && response.status === 200) {
                const unreadMail = await response.json();
                setCount(unreadMail)
            }
            else {
                setCount(null);
            }
        }
    }

    React.useEffect(() => { getUnreadMailCountAsync() })

    if (count === undefined && authenticationHeaders) {
        return (
            <Shimmer width="50%" />
        );
    }
    else if (count > 0) {
        return (<h2>You have {count} unread mail :-(</h2>)
    } else if (count === 0) {
        return (<h2>You have no unread mail! :-)</h2>)
    } else if (count === null) {
        return (<h2>Could not check mail status :-/</h2>)
    } else {
        return <></>
    }
}