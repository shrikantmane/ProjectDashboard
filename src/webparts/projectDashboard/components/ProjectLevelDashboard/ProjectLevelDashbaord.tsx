import * as React from 'react';
import { IProjectLevelDashboardProps } from './IProjectLevelDashboardProps';
import { IProjectLevelDashboardState } from './IProjectLevelDashboardState';

export default class ProjectLevelDashboard extends React.Component<IProjectLevelDashboardProps, IProjectLevelDashboardState> {
  
    public render(): React.ReactElement<IProjectLevelDashboardProps> {
    return (
      <div>
        {/* <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable> */}
      </div>
    );
  }
}
