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


  componentDidMount() {
  // console.log('componentDidUpdate',this.props.projectRoleResponsibility);
      this.getRoleResponsibility(this.props.projectRoleResponsibility);
  }

  // componentWillReceiveProps(nextProps) {
  //   console.log('componentWillReceiveProps',this.props.projectRoleResponsibility);
  //   console.log("roleResponsibilityList", this.state.roleResponsibilityList);
  //   if (this.props.projectRoleResponsibility != nextProps.projectRoleResponsibility)
  //     this.getRoleResponsibility(nextProps.projectRoleResponsibility);
  // }

  private getRoleResponsibility(projectRoleResponsibility: string) {
    sp.web.lists.getByTitle(projectRoleResponsibility).items
      .select("Owner/ID", "Owner/Title", "Owner/Department", "Roles_Responsibility").expand("Owner")
      .get()
      .then((response: Array<RoleResponsibility>) => {
        console.log('Responsibility -', response);
        this.setState({ roleResponsibilityList: response });
      });
  }

  public render(): React.ReactElement<IProjectRoleResponsibilityProps> {
    return (
      <div>
        {this.state.roleResponsibilityList != null
          ? this.state.roleResponsibilityList.map((item, key) => {
            return (
              <div>
                <div>
                  <span >{item.Owner ? item.Owner.Title : ""}</span>
                </div>
                <div>
                  <span >{item.Roles_Responsibility}</span>
                </div>
              </div>
            )
          })
          : null}
      </div>
    );
  }
}
