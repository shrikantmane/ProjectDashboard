import * as React from 'react';
import "primereact/resources/primereact.min.css";
import "primeicons/primeicons.css";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from './ProjectManagement.module.scss';
import { IProjectManagementProps } from './IProjectManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import ProjectListTable from './ProjectList/ProjectListTable';
import ProjectViewDetails from './ViewProject/ProjectViewDetails';
import { Switch, Route } from 'react-router-dom';

export default class ProjectManagement extends React.Component<IProjectManagementProps, {}> {
  public render(): React.ReactElement<IProjectManagementProps> {
    return (
      // <div>
      //   <ProjectListTable></ProjectListTable>
      // </div>
      <Switch>
        <Route exact path='/' component={ProjectListTable} />
        <Route path='/viewProjectDetails/:id' component={ProjectViewDetails} />
      </Switch>
    );
  }
}
