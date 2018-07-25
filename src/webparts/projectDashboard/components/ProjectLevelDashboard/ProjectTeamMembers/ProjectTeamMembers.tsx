import * as React from 'react';
import { IProjectTeamMembersProps } from './IProjectTeamMembersProps';
import { IProjectTeamMembersState } from './IProjectTeamMembersState';

export default class ProjectTeamMembers extends React.Component<IProjectTeamMembersProps, IProjectTeamMembersState> {
  
    public render(): React.ReactElement<IProjectTeamMembersProps> {
    return (
      <div>
        {/* <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable> */}
      </div>
    );
  }
}
