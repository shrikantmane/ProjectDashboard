import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectRoleResponsibilityProps } from './IProjectRoleResponsibilityProps';
import { IProjectRoleResponsibilityState } from './IProjectRoleResponsibilityState';
import { RoleResponsibility } from '../Project';

export default class ProjectProjectRoleResponsibility extends React.Component<IProjectRoleResponsibilityProps, IProjectRoleResponsibilityState> {

  constructor(props) {
    super(props);
    this.state = {
      roleResponsibilityList: new Array<RoleResponsibility>()
    };
  }


  componentWillReceiveProps(nextProps) {
    if (this.props.projectRoleResponsibility != nextProps.projectRoleResponsibility)
      this.getRoleResponsibility(nextProps.projectRoleResponsibility);
  }

  private getRoleResponsibility(projectRoleResponsibility: string) {
    sp.web.lists.getByTitle(projectRoleResponsibility).items
      .select("Owner/ID", "Owner/Title", "Owner/EMail", "Roles_Responsibility").expand("Owner")
      .get()
      .then((response: Array<RoleResponsibility>) => {
        let currentScope = this;
        let count = 1;
        response.forEach(item => {
          if (item.Owner) {
            let loginName = "i:0#.f|membership|" + item.Owner.EMail;
            sp.profiles.getPropertiesFor(loginName).then(function (result) {
              item.Owner.Department = result.UserProfileProperties[13].Value;
              //item.Owner.PictureURL = result.UserProfileProperties[16].Value;
              item.Owner.PictureURL = "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
              item.Owner.EMail +
              "&UA=0&size=HR64x64&sc=1531997060853";
              item.Owner.JobTitle = result.UserProfileProperties[21].Value;
              if (count == response.length) {
                currentScope.setState({ roleResponsibilityList: response });
              }
              count++;
            });
          } else {
            if (count == response.length) {
              currentScope.setState({ roleResponsibilityList: response });
            }
            count++;
          }
        });
      });
  }


  public render(): React.ReactElement<IProjectRoleResponsibilityProps> {
    return (
      <div className="col-xs-12 col-sm-5 rolesListPadding">
        <div className="well recommendedProjects userFeedback">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <h5>Roles and Responsibility</h5>
            </div>
            <div className="col-sm-12">
              <div className="profileDetails-container">
                {this.state.roleResponsibilityList != null
                  ? this.state.roleResponsibilityList.map((item, key) => {
                    let email = "https://uatalpha-my.sharepoint.com/_layouts/15/me.aspx/?p=" + item.Owner.EMail + "&v=work";
                    return (
                      <div className="row">
                        <div className="col-sm-12">
                          <div className="row">
                            <div className="col-sm-2">
                              <img className="img-responsive image-style rolesImage" src={item.Owner.PictureURL} alt="" />
                            </div>
                            <div className="col-sm-2">
                              <div className="profileDetail">
                                <div className="profileName">
                                  <a href={email} target="_blank">{item.Owner ? item.Owner.Title : ""}</a>
                                </div>
                                <h5 className="deptName">{item.Owner ? item.Owner.Department : ""}</h5>
                              </div>
                            </div>
                            <div className="col-sm-8 float-left">
                            <ul className="profileRoles">
                                  <li>{item.Roles_Responsibility}</li>
                                </ul>
                            </div>
                          </div>
                        </div>
                      </div>
                    )
                  })
                  : null}
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
