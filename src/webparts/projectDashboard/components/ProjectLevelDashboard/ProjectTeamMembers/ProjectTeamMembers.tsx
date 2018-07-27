import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectTeamMembersProps } from './IProjectTeamMembersProps';
import { IProjectTeamMembersState } from './IProjectTeamMembersState';
import { TeamMember } from '../Project';

export default class ProjectTeamMembers extends React.Component<IProjectTeamMembersProps, IProjectTeamMembersState> {

  constructor(props) {
    super(props);
    this.state = {
      teamMemberList: new Array<TeamMember>()
    };
  }

  componentWillReceiveProps(nextProps) {
    if (this.props.projectTeamMembers != nextProps.projectTeamMembers)
      this.getTeamMember(nextProps.projectTeamMembers);
  }

  private getTeamMember(projectTeamMembers: string) {
    sp.web.lists.getByTitle(projectTeamMembers).items
      .select("Team_x0020_Member/ID", "Team_x0020_Member/Title", "Team_x0020_Member/EMail", "Team_x0020_Member/Department").expand("Team_x0020_Member")
      .get()
      .then((response: Array<TeamMember>) => {
        console.log('Team Member -', response);
        this.setState({teamMemberList : response});
      });

  }

  public render(): React.ReactElement<IProjectTeamMembersProps> {
    return (
      <div>
        {this.state.teamMemberList != null
          ? this.state.teamMemberList.map((item, key) => {
            if (item.Team_x0020_Member) {
              return (
                <div>
                   <div>
                  <span >{item.Team_x0020_Member.Title}</span>
                  </div>
                  <div>
                  <span >{item.Team_x0020_Member.EMail}</span>
                  </div>
                </div>
              )
            }
          })
          : null}
      </div>
    );
  }
}
