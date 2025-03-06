// <copyright file="router.js" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

import * as React from "react";
import {
    useNavigate,
    useLocation,
} from 'react-router-dom';
import * as microsoftTeams from "@microsoft/teams-js";
import AppCacheTab from "../components/app-cache-tab";
import SecondCachedTab from "../components/app-cache-tab-2";

export const Tab = React.memo((props: {viewSelection: "tab1" | "tab2"}) => {
    const [entityId, setEntityId] = React.useState<string>("");
    const [appInitialized, setAppInitialized] = React.useState(false);
    const navigate = useNavigate();
    const viewSelection = props.viewSelection || "tab1";
    const location = useLocation();

    React.useEffect(() => {
        if (location.state && location.state.entityId) {
            console.log(">>>Tab updating entityId");
            setEntityId(location.state.entityId);
        } else {
            const params = new URLSearchParams(window.location.search);
            const entityIdParam = params.get("entityId");
            if (entityIdParam) {
                setEntityId(entityIdParam);
            }
        }
    }, [location]);

    React.useEffect(() => {
        // Initialize the Microsoft Teams SDK
        const app = microsoftTeams.app;

        if (!appInitialized) {
            console.log("Initializing TeamsJS SDK");

            app.initialize().then(() => {
                // Check if the framecontext is a cacheable one
                if (window.location.pathname === "/appCacheTab" || window.location.pathname === "/appCacheTab2") {

                    microsoftTeams.teamsCore.registerBeforeUnloadHandler((readyToUnload: any) => {
                        readyToUnload();
                        console.log("sending readyToUnload to TEAMS");
                        return true;
                    });

                    microsoftTeams.teamsCore.registerOnLoadHandler((data: any) => {
                        if (data.entityId) {
                            console.log("Load handler sending new entityId to TEAMS " + data.entityId);
                            if (data.entityId === "second") {
                                navigate("/appCacheTab2?entityId=" + data.entityId, {state: {entityId: data.entityId}});
                            } else {
                                navigate("/appCacheTab?entityId=" + data.entityId, {state: {entityId: data.entityId}});
                            }
                        }
                    });
                }

                setAppInitialized(true);
            }).catch(function (error: any) {
                console.error(error, "Could not initialize TeamsJS SDK.");
            });
        }

        return () => {
            console.log("useEffect cleanup - Tab");
        };
    }, [navigate, appInitialized]);

    return (
        <React.Fragment>
            { appInitialized ? (
                <>
                { viewSelection === "tab1" ? <AppCacheTab entityId={entityId} /> : <SecondCachedTab entityId={entityId} />}
                </> ) : null
            }
        </React.Fragment>
    );
});
