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
      .select("Team_x0020_Member/ID", "Team_x0020_Member/Title", "Team_x0020_Member/EMail").expand("Team_x0020_Member")
      .get()
      .then((response: Array<TeamMember>) => {
        let currentScope = this;
        let count = 1;
        response.forEach(item => {
          if (item.Team_x0020_Member) {
            let loginName = "i:0#.f|membership|" + item.Team_x0020_Member.EMail;
            sp.profiles.getPropertiesFor(loginName).then(function (result) {
              item.Team_x0020_Member.Department = result.UserProfileProperties[13].Value;
              //item.Team_x0020_Member.PictureURL = result.UserProfileProperties[16].Value;
              item.Team_x0020_Member.PictureURL = "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
              item.Team_x0020_Member.EMail +
              "&UA=0&size=HR64x64&sc=1531997060853";
              item.Team_x0020_Member.JobTitle = result.UserProfileProperties[22].Value;
              if (count == response.length) {
                currentScope.setState({ teamMemberList: response });
              }
              count++;
            });
          } else {
            if (count == response.length) {
              currentScope.setState({ teamMemberList: response });
            }
            count++;
          }
        });
      });

  }

  public render(): React.ReactElement<IProjectTeamMembersProps> {
    return (
      <div className="col-xs-12 col-sm-3 teamMemeberListPadding">
        <div className="well recommendedProjects userFeedback">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <h5>Team Members</h5>
            </div>
            <div className="col-sm-12">
              <div className="profileDetails-container">
                {this.state.teamMemberList != null ?
                  this.state.teamMemberList.map((item, key) => {
                    if (item.Team_x0020_Member) {
                      let email = "https://uatalpha-my.sharepoint.com/_layouts/15/me.aspx/?p=" + item.Team_x0020_Member.EMail + "&v=work";
                      return (<div className="row">
                        <div className="col-sm-12">
                          <div className="row">
                            <div className="col-sm-3">
                              <img className="img-responsive image-style" src={item.Team_x0020_Member.PictureURL} alt="" />
                            </div>
                            <div className="col-sm-9">
                              <div className="profileDetail">
                                <div className="profileName">
                                  <a href={email} target="_blank">{item.Team_x0020_Member.Title}</a>
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
