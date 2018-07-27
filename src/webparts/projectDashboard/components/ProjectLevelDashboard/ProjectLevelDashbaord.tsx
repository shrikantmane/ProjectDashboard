import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectLevelDashboardProps } from "./IProjectLevelDashboardProps";
import { IProjectLevelDashboardState } from "./IProjectLevelDashboardState";
import { Project, Tag } from "./Project";
import ProjectMildstone from "./ProjectMildstone/ProjectMildstone";
import ProjectPlan from "./ProjectPlan/ProjectPlan";
import ProjectTaskList from "./ProjectTaskList/ProjectTaskList";
import ProjectDocument from "./ProjectDocument/ProjectDocument";
import ProjectTeamMembers from "./ProjectTeamMembers/ProjectTeamMembers";
import ProjectTeamConversation from "./ProjectTeamConversation/ProjectTeamConversation";
import ProjectProjectRoleResponsibility from "./ProjectRoleResponsibility/ProjectRoleResponsibility";
import styles from "../ProjectDashboard.module.scss";

export default class ProjectLevelDashboard extends React.Component<
  IProjectLevelDashboardProps,
  IProjectLevelDashboardState
  > {
  constructor(props) {
    super(props);
    this.state = {
      project: new Project(),
      tagList: new Array<Tag>()
    };
  }
  componentDidMount() {
    const {
      match: { params }
    } = this.props;
    this.getProjectByProjectID(params.id);
    this.getProjectTagsByProjectID(params.id);
  }

  private getProjectByProjectID(id: number) {
    let project = new Project();
    let filter = "ID eq " + id;
    sp.web.lists
      .getByTitle("Project")
      .items.select(
        "Project",
        "DueDate",
        "Priority",
        "Task_x0020_List",
        "Schedule_x0020_List",
        "Project_x0020_Document",
        "Project_x0020_Team_x0020_Members"
      )
      .filter(filter)
      .getAll()
      .then((response: Array<Project>) => {
        console.log("Project items", response);
        if (response && response.length > 0) {
          project.Project = response[0].Project;
          project.DueDate = response[0].DueDate;
          project.Priority = response[0].Priority;
          project.Task_x0020_List = response[0].Task_x0020_List;
          project.Schedule_x0020_List = response[0].Schedule_x0020_List;
          project.Project_x0020_Document = response[0].Project_x0020_Document;
          project.Project_x0020_Team_x0020_Members =
            response[0].Project_x0020_Team_x0020_Members;
          this.setState({ project: project });
        }
      })
      .catch((e: Error) => {
        console.log(`There was an error : ${e.message}`);
      });
  }

  private getProjectTagsByProjectID(id: number) {
    let filter = "Project/ID eq " + id;
    sp.web.lists.getByTitle("Project Tags").items
      .select("Tag","Color")
      .filter(filter)
      .get()
      .then((response : Array<Tag>) => {
        console.log("tagList", response);
       this.setState({tagList : response})
      });

  }
  public render(): React.ReactElement<IProjectLevelDashboardProps> {
    return (
      <div className="ProjectLevelDashboard">
        <div className="container-fluid">
          <section className="main-content-section dashboardSection">
            <div className="wrapper">
              <div className="row conversationTasks">
                <div className="project-tabs col-xs-12 col-sm-12 col-md-12 col-lg-12">
                  <div className="projectName">{this.state.project.Project}
                <span className="due-date-style">{ new Date(this.state.project.DueDate).toDateString()}</span>
                    <div className="tagList">
                      <span className="delayedStatus priority-btn">{this.state.project.Priority}</span>
                    </div>
                  </div>
                </div>
                <div className="col-lg-12 col-md-12 col-sm-12">
                  <div className="well recommendedProjects">
                    <div className="row">
                      <div className="Status-block col-xs-12 col-sm-12 col-md-12 col-lg-12">
                        <div className="milestoneHeader">
                          <div className="row">
                            <div className="col-md-6 col-12">
                              <div className="tagList">
                              { this.state.tagList ? this.state.tagList.map((item, key) => {
                                 return  (<span className="pinkTag" style={{backgroundColor : item.Color}}>{item.Tag}</span>)

                              }) : null}                               
                              </div>
                            </div>
                          </div>
                        </div>
                        <ProjectMildstone scheduleList={this.state.project.Schedule_x0020_List}></ProjectMildstone>                      
                      </div>
                    </div>
                  </div>
                </div>

                  <ProjectTaskList scheduleList={this.state.project.Schedule_x0020_List}></ProjectTaskList>
                  <ProjectDocument projectDocument={this.state.project.Project_x0020_Document}></ProjectDocument>
                  <div className="clearfix"></div>
                  <ProjectTeamMembers projectTeamMembers={this.state.project.Project_x0020_Team_x0020_Members}></ProjectTeamMembers>
                  <ProjectProjectRoleResponsibility projectRoleResponsibility ={"Project Information"} ></ProjectProjectRoleResponsibility>

                {/* row conversationTasks */}
              </div>
            </div>
          </section>
        </div>
      </div>
    );
  }
}
