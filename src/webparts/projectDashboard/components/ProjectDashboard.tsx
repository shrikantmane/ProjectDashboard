import * as React from 'react';
import { Web, sp, ItemAddResult } from '@pnp/sp';
import 'primereact/resources/primereact.min.css';
import 'primeicons/primeicons.css';
import 'bootstrap/dist/css/bootstrap.min.css';

import { IProjectDashboardProps } from './IProjectDashboardProps';
import { IProjectDashboardState } from './IProjectDashboardState';
import CEODashboard from './CEODashboard/CEODashbaord';
import DepartmentHeadDashboard from '../components/DepartmentHeadDashboard/DepartmentHeadDashboard';
import { UserType }  from "./ProjectUser";
export default class ProjectDashboard extends React.Component<IProjectDashboardProps, IProjectDashboardState> {

  constructor(props) {
    super(props);
    this.state = {
      userType: UserType.Unknow
    };
  }

  componentDidMount() {
    this.getCurrentUser();
  }

  public render(): React.ReactElement<IProjectDashboardProps> {   

    let dashboard;
    switch(this.state.userType) {
        case UserType.CEO:
        dashboard = (<CEODashboard></CEODashboard>);
            break;
        case UserType.DepartmentHead:
        dashboard = (<DepartmentHeadDashboard></DepartmentHeadDashboard>);
            break;
        default:
        dashboard = (<CEODashboard></CEODashboard>);
    }
    return (
      <div>
          { dashboard }
      </div>
    );
  }

  // get current user details
   private getCurrentUser() :void {
      sp.web.currentUser.get().then(result => {
        this.getUserGroup(result.Id);
      });
    }

    // get user group
    getUserGroup(Id){
      sp.web.siteUsers.getById(Id).groups.get().then(res => {
      let userType;
         if(res && res.length > 0){
            if(res.length > 1){
              userType = UserType.Admin;
            }else{
              if(res[0].LoginName == "CEO_COO" ){
                userType = UserType.CEO;
              }else if(res[0].LoginName == "Department Head"){
                userType = UserType.DepartmentHead;
              }
            }
            this.setState({ userType: userType})
         }
      })
    }
}
