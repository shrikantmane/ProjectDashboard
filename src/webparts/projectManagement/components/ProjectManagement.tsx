import * as React from 'react';
import "primereact/resources/primereact.min.css";
import "primeicons/primeicons.css";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from './ProjectManagement.module.scss';
import { IProjectManagementProps } from './IProjectManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import ProjectListTable from './ProjectList/ProjectListTable';
import TeamListTable from './TeamMembers/TeamListTable';
import RiskListTable from './ProjectRisk/RiskListTable';
import InformationListTable from './ProjectInformation/InformationListTable'
import RequirementListTable from './ProjectRequirement/RequirementListTable'
import DocumentListTable from './ProjectDocuments/DocumentListTable'
import ScheduleListTable from './ProjectSchedule/ScheduleListTable'
import FullCalendar from 'fullcalendar-reactwrapper';
import { SPComponentLoader } from "@microsoft/sp-loader";
import ProjectViewDetails from './ViewProject/ProjectViewDetails';
import { Switch, Route } from 'react-router-dom';
import { Web, sp, ItemAddResult } from "@pnp/sp";
import { UserType } from "./ProjectUser";
import { find, filter, sortBy } from "lodash";

export default class ProjectManagement extends React.Component<IProjectManagementProps, {
  userType: UserType;
  department: string;
}> {
  constructor(props) {
    super(props);
    this.state = {
      userType: UserType.Unknow,
      department: ""
    };
  }
  componentDidMount() {
    SPComponentLoader.loadCss(
      "https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.9.0/fullcalendar.css"
    );
    this.getCurrentUser();


  }
  // get current user details
  private getCurrentUser(): void {
    sp.web.currentUser.get().then(result => {
      this.getUserGroup(result);
    });
  }
  getUserGroup(result) {
    sp.web.siteUsers
      .getById(result.Id)
      .groups.get()
      .then(res => {
        let userType;
        if (res && res.length > 0) {
          let adminData = find(res, function (o: any) { return o.LoginName === 'hbc Admin'; });
          if (adminData) {
            this.setState({ userType: UserType.Admin });
          } else {
            for (let i = 0; i < res.length; i++) {
              if (res[i].LoginName == "CEO_COO" || res[i].LoginName == "Members") {
                this.setState({ userType: UserType.CEO });
                break;
              }
            }
          }
        }
        console.log(this.state.userType);
      });
  }
  public render(): React.ReactElement<IProjectManagementProps> {

    // return (
    //   <div className={ styles.projectManagement }>
    //     <div className={ styles.container }>
    //       <div className={ styles.row }>
    //         <div className={ styles.column }>
    //           <span className={ styles.title }>Welcome to SharePoint!</span>
    //           <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
    //           <p className={ styles.description }>{escape(this.props.description)}</p>
    //           <a href="https://aka.ms/spfx" className={ styles.button }>
    //             <span className={ styles.label }>Learn more</span>
    //           </a>
    //         </div>
    //       </div>
    //     </div>
    //   </div>
    // );
    return (
      this.state.userType === 'CEO' ? <div>Access Denied</div> :
        <Switch>
          <Route exact path='/' component={ProjectListTable} />
          <Route path='/viewProjectDetails/:id' component={ProjectViewDetails} />
        </Switch>

    );
  }
}
