import * as React from "react";
import { Web, sp, ItemAddResult } from "@pnp/sp";
import "primereact/resources/primereact.min.css";
import "primeicons/primeicons.css";
import "bootstrap/dist/css/bootstrap.min.css";
import { find } from "lodash";

import { IMainDashboardProps } from "./IMainDashboardProps";
import { IMainDashboardState } from "./IMainDashboardState";
import CEODashboard from "../CEODashboard/CEODashbaord";
import DepartmentHeadDashboard from "../DepartmentHeadDashboard/DepartmentHeadDashboard";
import TeamMemberDashboard from "../TeamMemberDashboard/TeamMemberDashboard";
import { UserType } from "./ProjectUser";
export default class MainDashboard extends React.Component<
  IMainDashboardProps,
  IMainDashboardState
  > {
  constructor(props) {
    super(props);
    this.state = {
      userType: UserType.Unknow,
      department: ""
    };
  }

  componentDidMount() {
    this.getCurrentUser();
  }

  public render(): React.ReactElement<IMainDashboardProps> {
    let dashboard = null;
    if (this.state.userType != UserType.Unknow) {
      switch (this.state.userType) {
        case UserType.CEO:
          dashboard = <CEODashboard webPartTitle={this.props.webPartTitle} {...this.props} />;
          break;
        case UserType.DepartmentHead:
          dashboard = <DepartmentHeadDashboard department={this.state.department} />;
          break;
        case UserType.Members:
          dashboard = <TeamMemberDashboard />;
          break;
        default:
          dashboard = <h6>No Role Found</h6>;
          break;
      }
    }
    return <div>{dashboard}</div>;
  }

  // get current user details
  private getCurrentUser(): void {
    sp.web.currentUser.get().then(result => {
      this.getUserGroup(result);
    });
  }

  // get user group
  getUserGroup(result) {
    sp.web.siteUsers
      .getById(result.Id)
      .groups.get()
      .then(res => {
        let userType;
        if (res && res.length > 0) {
          let ceo = "";
          let dep = "";
          let team = "";
          res.forEach(item => {
            if (item.LoginName == "CEO_COO" || item.LoginName == "Admin") {
              ceo = "CEO_COO";
            }
            if (item.LoginName == "Department Head") {
              dep = "Department Head";
            }
            if (item.LoginName == "Members") {
              team = "Members";
            }
          });

          if (ceo == "CEO_COO") {
            this.setState({ userType: UserType.CEO });
          } else if (dep == "Department Head") {
            this.getDepartment(result.Email);
          } else if ((team == "Members"))
            this.setState({ userType: UserType.Members });
        }
      });
  }

  private getDepartment(email) {
    let currentScope = this;
    let loginName = "i:0#.f|membership|" + email;
    sp.profiles.getPropertiesFor(loginName).then(function (result) {
      let department = result.UserProfileProperties ? result.UserProfileProperties[13].Value : "";
      currentScope.setState({ userType: UserType.DepartmentHead, department: department });
    }).catch(function (err) {
      console.log("Error: " + err);
    });
  }
}
