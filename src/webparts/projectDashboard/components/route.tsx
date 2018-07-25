import * as React from 'react';
import { Switch, Route } from 'react-router-dom';
import ProjectDashboard from "./ProjectDashboard";
import ProjectLevelDashboard from "../components/ProjectLevelDashboard/ProjectLevelDashbaord";

export default class RouteComponent extends React.Component<any, any> {
  
    public render(): React.ReactElement<any> {
    return (
        <Switch>
            <Route exact path='/' component={ProjectDashboard}/>
            <Route path='/project' component={ProjectLevelDashboard}/>
        </Switch>
    );
  }
}
