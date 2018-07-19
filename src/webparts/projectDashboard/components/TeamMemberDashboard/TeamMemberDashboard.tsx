import * as React from 'react';
import { ITeamMemberDashboardProps } from './ITeamMemberDashboardProps';
import { ITeamMemberDashboardState } from './ITeamMemberDashboardState';
import CEOProjectTable from '../CEOProjectsTable/CEOProjectTable';
import CEOProjectTimeLine from '../CEOProjectTimeLine/CEOProjectTimeLine';

export default class TeamMemberDashboard extends React.Component<ITeamMemberDashboardProps, ITeamMemberDashboardState> {
  
    public render(): React.ReactElement<ITeamMemberDashboardProps> {
    return (
      <div>
         <h6>Team Member Dashboard</h6>
      </div>
    );
  }
}
