import * as React from 'react';
import { IDepartmentHeadDashboardProps } from './IDepartmentHeadDashboardProps';
import { IDepartmentHeadDashboardState } from './IDepartmentHeadDashboardState';
import ProjectTaskList from "./ProjectTaskList/ProjectTaskList";
import ProjectTeamMembers from "./ProjectTeamMembers/ProjectTeamMembers";
import ProjectTeamConversation from "./ProjectTeamConversation/ProjectTeamConversation";
import CEOProjectInformation from "./CEOProjectInformation/CEOProjectInformation";
import DepartmentProjectInformation from "./DepartmentProjectInformation/DepartmentProjectInformation";

export default class DepartmentHeadDashboard extends React.Component<IDepartmentHeadDashboardProps, IDepartmentHeadDashboardState> {

  public render(): React.ReactElement<IDepartmentHeadDashboardProps> {
    return (
      <div className="ProjectLevelDashboard">
        <div className="container-fluid">
          <section className="main-content-section dashboardSection">
            <div className="wrapper">
              <div className="row conversationTasks">
                <div className="col-md-12 col-xs-12 cardPadding">
                  <CEOProjectInformation></CEOProjectInformation>
                </div>
                <div className="col-md-12 col-xs-12 cardPadding">
                  <DepartmentProjectInformation department={this.props.department}></DepartmentProjectInformation>
                </div>
                <div className="col-md-12 col-xs-12 cardPadding">
                  <ProjectTaskList department={this.props.department}></ProjectTaskList>
                </div>
                <div className="col-md-3 col-xs-12 cardPadding">
                  <ProjectTeamMembers department={this.props.department}></ProjectTeamMembers>
                </div>
                <div className="col-md-9 col-xs-12 cardPadding">
                  <ProjectTeamConversation ></ProjectTeamConversation>
                </div>
              </div>
            </div>
          </section>
        </div>
      </div>
    );
  }
}
