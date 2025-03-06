// <copyright file="app-cache-tab.tsx" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

import React from "react";
import "../components/index.css";
import { app } from "@microsoft/teams-js";

const supportedTabEntities = ["second"];

const SecondCachedTab = React.memo((props: {entityId: string}) => {
    const {entityId} = props;
    const isSupportedView = React.useMemo(() => supportedTabEntities.includes(entityId), [entityId]);
    
    React.useEffect(()=>{
        if (isSupportedView) {
            setTimeout(() => {
            console.log(">>>SecondCachedTab sending notifySuccess");
            app.notifySuccess();
            }, 100);
        }
      }, [isSupportedView]);
       return isSupportedView ? (
        <div>
            <h2 style={{color: 'red', marginTop: "120px" , backgroundColor: 'yellow'}}>Second Component</h2>
        </div>) : null;
});

export default SecondCachedTab;