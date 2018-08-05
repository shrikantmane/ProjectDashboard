import * as React from 'react';
import { ITeamMemberDashboardProps } from './ITeamMemberDashboardProps';
import { ITeamMemberDashboardState } from './ITeamMemberDashboardState';

export default class TeamMemberDashboard extends React.Component<ITeamMemberDashboardProps, ITeamMemberDashboardState> {
  
    public render(): React.ReactElement<ITeamMemberDashboardProps> {
    return (
      <div className="col-xs-12 col-sm-9">
        <div className="well recommendedProjects userFeedback">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <h5>Roles and Responsibility</h5>
            </div>
            <div className="col-sm-12">
              
            </div>
          </div>
        </div>
      </div>
    );
  }
}
