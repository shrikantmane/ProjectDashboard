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
import { UserType } from "./MainDashboard/ProjectUser";

import { Switch, Route } from 'react-router-dom';
export default class ProjectDashboard extends React.Component<IProjectDashboardProps, IProjectDashboardState> {
    constructor(props) {
        super(props);
        this.state = {
            userType: UserType.Unknow,
            department: ""
        };
    }
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getCurrentUser();
    }
    // get current user details
    private getCurrentUser(): void {
        sp.web.currentUser.get().then(result => {
            this.getUserGroup(result);
        });
    }
    getUserGroup(result) {
        sp.web.siteUsers
            .getById(result.Id)
            .groups.get()
            .then(res => {
                let userType;
                if (res && res.length > 0) {

                    for (let i = 0; i < res.length; i++) {
                        if (res[i].LoginName == "CEO_COO" || res[i].LoginName == "Department Head" || res[i].LoginName == "hbc Admin") {
                            this.setState({ userType: UserType.CEO });
                            break;
                        }
                    }
                }
                console.log(this.state.userType);
            });
    }
    public render(): React.ReactElement<IProjectDashboardProps> {
        return (
            this.state.userType === UserType.CEO ?
                <Switch>
                    <Route exact path='/' component={MainDashboard} />
                    <Route path='/projectDetails/:id' component={ProjectLevelDashboard} />
                </Switch>
                : <div>Access Denied</div>
        );
    }
}

