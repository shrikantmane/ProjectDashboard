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
      let currentScope = this;
      let count = 1;
      response.forEach(item => {
        let loginName = "i:0#.f|membership|" + item.Team_x0020_Member.EMail;
        sp.profiles.getPropertiesFor(loginName).then(function (result) {
          item.Team_x0020_Member.Department = result.UserProfileProperties[13].Value;
          item.Team_x0020_Member.PictureURL = result.UserProfileProperties[16].Value;
          item.Team_x0020_Member.JobTitle = result.UserProfileProperties[21].Value;
          if (count == response.length) {
            currentScope.setState({ teamMemberList: response });
          }
          count ++;
        });
      });
      });

  }

  public render(): React.ReactElement<IProjectTeamMembersProps> {
    return (
      <div className="col-xs-12 col-sm-3">
        <div className="well recommendedProjects userFeedback">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <h5>Team Members</h5>
            </div>
            <div className="col-sm-12">
              <div className="profileDetails-container">
                {this.state.teamMemberList != null ?
                  this.state.teamMemberList.map((item, key) => {
                    if(item.Team_x0020_Member){
                    return (<div className="row">
                      <div className="col-sm-12">
                        <div className="row">
                          <div className="col-sm-2">
                            <img className="img-responsive image-style" src={ item.Team_x0020_Member.PictureURL} alt="" />
                          </div>
                          <div className="col-sm-8">
                            <div className="profileDetail">
                              <div className="profileName">
                                <h4>{item.Team_x0020_Member.Title}</h4>
                              </div>
                              <div className="profileDesignation">
                                <span className="designationTag">{item.Team_x0020_Member.JobTitle}</span>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>)
                    }
                  }) : null
                }
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
