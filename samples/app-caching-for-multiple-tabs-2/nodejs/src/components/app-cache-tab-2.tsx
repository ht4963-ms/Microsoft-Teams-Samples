// <copyright file="app-cache-tab.tsx" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

import React from "react";
import { useParams } from 'react-router-dom';
import "./index.css";
import { app } from "@microsoft/teams-js";
import { loadNewEntityData } from "./utils";

export const AppCacheTab2 = () => {

    const { entityId } = useParams<{ entityId: string }>();
    const [currentEntityId, setCurrentEntityId] = React.useState<string>("");
    const [loading, setLoading] = React.useState<boolean>(true);

    React.useEffect(() => {
        if (entityId !== currentEntityId) {
            console.log(`Entity ID changed from ${currentEntityId} to ${entityId}`);
            setCurrentEntityId(entityId || "");
            loadNewEntityData(entityId || "", setLoading);
        }
    }, [entityId]);
    
    React.useEffect(() => {
        console.log(`Page 2 sending notifySuccess`);
        app.notifySuccess();
    }, []);
    
    return (
        <div style={{ backgroundColor: 'yellow', color: 'red', height: '800px', padding: '20px' }}>
            {loading ? <div>Loading...</div> :
                (
                    <div>
                        <h2>Page 2</h2>
                        <h3>Entity ID: {entityId}</h3>
                    </div>
                )
            }
            </div>
    );
};


