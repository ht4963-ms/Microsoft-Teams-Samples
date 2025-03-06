// <copyright file="router.js" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// </copyright>

import * as React from "react";
import {
    BrowserRouter,
    Route,
    Routes,
} from 'react-router-dom';
import Configure from "../components/configure";
import Index from "../components/index";
import { Tab } from "../components/tab";

export const AppRoute = () => {

    return (
        <React.Fragment>
            <BrowserRouter>
                <Routes>
                    <Route path="/" element={<Index />} />
                    <Route path="/configure" element={<Configure />}/>
                    <Route path="/appCacheTab" element={<Tab viewSelection="tab1" />}/>
                    <Route path="/appCacheTab2" element={<Tab viewSelection="tab2" />}/>
                </Routes>
            </BrowserRouter>
        </React.Fragment>
    );
};