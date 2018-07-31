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
        console.log('Responsibility -', response);
        let currentScope = this;
        let count = 1;
        response.forEach(item => {
          if (item.Owner) {
            let loginName = "i:0#.f|membership|" + item.Owner.EMail;
            sp.profiles.getPropertiesFor(loginName).then(function (result) {
              item.Owner.Department = result.UserProfileProperties[13].Value;
              item.Owner.PictureURL = result.UserProfileProperties[16].Value;
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
      <div className="col-xs-12 col-sm-4">
        <div className="well recommendedProjects userFeedback">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <h5>Roles and Responsibility</h5>
            </div>
            <div className="col-sm-12">
              <div className="profileDetails-container">
                {this.state.roleResponsibilityList != null
                  ? this.state.roleResponsibilityList.map((item, key) => {
                    return (
                      <div className="row">
                        <div className="col-sm-12">
                          <div className="row">
                            <div className="col-sm-2">
                              <img className="img-responsive image-style" src={item.Owner.PictureURL} alt="" />
                            </div>
                            <div className="col-sm-5">
                              <div className="profileDetail">
                                <div className="profileName">
                                  <h4>{item.Owner ? item.Owner.Title : ""}</h4>
                                </div>
                                <ul className="profileRoles">
                                  <li>{item.Roles_Responsibility}</li>
                                </ul>
                              </div>
                            </div>
                            <div className="col-sm-3 pull-right">
                              <h5 className="deptName">{item.Owner ? item.Owner.Department : ""}</h5>
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
