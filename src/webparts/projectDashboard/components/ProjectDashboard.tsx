import * as React from 'react';
import 'primereact/resources/primereact.min.css';
import 'primeicons/primeicons.css';
import 'bootstrap/dist/css/bootstrap.min.css';

import { IProjectDashboardProps } from './IProjectDashboardProps';
import CEODashboard from './CEODashboard/CEODashbaord';
export default class ProjectDashboard extends React.Component<IProjectDashboardProps, {}> {
  public render(): React.ReactElement<IProjectDashboardProps> {
    return (

      <div>
        <CEODashboard></CEODashboard>
      </div>
    );
  }
}
