import * as React from "react";
import { Web, sp, ItemAddResult } from "@pnp/sp";
import "primereact/resources/primereact.min.css";
import "primeicons/primeicons.css";
import "bootstrap/dist/css/bootstrap.min.css";
import { IProjectDashboardProps } from "./IProjectDashboardProps";
import { IProjectDashboardState } from "./IProjectDashboardState";
import MainDashboard from "./MainDashboard/MainDashboard";
import ProjectLevelDashboard from "../components/ProjectLevelDashboard/ProjectLevelDashbaord";
import { SPComponentLoader } from "@microsoft/sp-loader";

import { Switch, Route } from 'react-router-dom';
export default class ProjectDashboard extends React.Component<IProjectDashboardProps, IProjectDashboardState> {

    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
    }

    public render(): React.ReactElement<IProjectDashboardProps> {
        return (
            <Switch>
                <Route exact path='/' component={MainDashboard} />
                <Route path='/projectDetails/:id' component={ProjectLevelDashboard} />
            </Switch>
        );
    }
}

