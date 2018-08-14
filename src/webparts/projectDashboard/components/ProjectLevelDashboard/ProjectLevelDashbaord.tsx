import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import moment from 'moment/src/moment';
import { IProjectLevelDashboardProps } from "./IProjectLevelDashboardProps";
import { IProjectLevelDashboardState } from "./IProjectLevelDashboardState";
import { Project, Tag, File } from "./Project";
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
      tagList: new Array<Tag>(),
      attachment: new File()
    };
  }

  componentDidMount() {
    const {
      match: { params }
    } = this.props;
    this.getProjectByProjectID(params.id);
    this.getProjectTagsByProjectID(params.id);
    var elmnt = document.getElementById('projectDashboard_projectName');
    if (elmnt)
      elmnt.scrollIntoView();
  }

  private getProjectByProjectID(id: number) {
    let project = new Project();
    let file = new File();
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
      "Project_x0020_Team_x0020_Members",
      "Project_x0020_Infromation",
      "Project_x0020_Comments",
      "Task_x0020_Comments",
      "AttachmentFiles",
      "AttachmentFiles/ServerRelativeUrl",
      "AttachmentFiles/FileName"
      )
      .expand("AttachmentFiles")
      .filter(filter)
      .getAll()
      .then((response: any) => {
        if (response && response.length > 0) {
          console.log('Attachment : ', response);
          project.Project = response[0].Project;
          project.DueDate = response[0].DueDate;
          project.Priority = response[0].Priority;
          project.Task_x0020_List = response[0].Task_x0020_List;
          project.Schedule_x0020_List = response[0].Schedule_x0020_List;
          project.Project_x0020_Document = response[0].Project_x0020_Document;
          project.Project_x0020_Team_x0020_Members = response[0].Project_x0020_Team_x0020_Members;
          project.Project_x0020_Infromation = response[0].Project_x0020_Infromation;
          project.Task_x0020_Comments = response[0].Task_x0020_Comments;
          file.Name = response[0].AttachmentFiles.length > 0 ? response[0].AttachmentFiles[0].FileName : '';
          file.ServerRelativeUrl = response[0].AttachmentFiles.length > 0 ? response[0].AttachmentFiles[0].ServerRelativeUrl : '';
          this.setState({ project: project, attachment: file });
        }
      })
      .catch((e: Error) => {
        console.log(`There was an error : ${e.message}`);
      });
  }

  private getProjectTagsByProjectID(id: number): void {
    let filterString = "Projects/ID eq " + id;
    sp.web.lists.getByTitle("Project Tags").items
      .select("Projects/ID", "Tag", "Color").expand("Projects")
      .filter(filterString)
      .get()
      .then((response) => {
        if (response.length > 0) {
          let tags = new Array<Tag>();
          response.forEach(element => {
            tags.push({
              Tag: element.Tag,
              Color: element.Color
            });
          });
          this.setState({ tagList: response });
        }
      });
  }
  public render(): React.ReactElement<IProjectLevelDashboardProps> {
    let projectOutlineButton = this.state.attachment.ServerRelativeUrl !== '' ? <a href={this.state.attachment.ServerRelativeUrl} target="_blank" style={{ float: 'right' }} className="btnoutlineproject btn btn-sm">Project Outline</a> : null
    return (
      <div className="ProjectLevelDashboard">
        <div className="container-fluid">
          <section className="main-content-section dashboardSection">
            <div className="wrapper">
              <div className="row conversationTasks">
                <div className="project-tabs col-xs-12 col-sm-12 col-md-12 col-lg-12 cardPadding">
                  <div className="projectName" id="projectDashboard_projectName">{this.state.project.Project}
                    <span className="due-date-style">Due on : {moment(this.state.project.DueDate).format("DD MMM YYYY")}</span>
                    <div className="tagList">
                      <span className="delayedStatus priority-btn">Priority: {this.state.project.Priority}</span>
                    </div>
                    {projectOutlineButton}
                  </div>
                </div>
                <div className="col-lg-12 col-md-12 col-sm-12 cardPadding">
                  <div className="card well recommendedProjects">
                    <div className="row">
                      <div className="Status-block col-xs-12 col-sm-12 col-md-12 col-lg-12">
                        <div className="milestoneHeader">
                          <div className="row">
                            <div className="col-md-6 col-12">
                              <div className="tagList">
                                {this.state.tagList ? this.state.tagList.map((item, key) => {
                                  return (<span className="pinkTag" style={{ backgroundColor: item.Color }}>{item.Tag}</span>)

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
                <ProjectPlan scheduleList={this.state.project.Schedule_x0020_List} commentList={this.state.project.Task_x0020_Comments} ></ProjectPlan>
                <ProjectTaskList taskList={this.state.project.Task_x0020_List}></ProjectTaskList>
                <ProjectDocument projectDocument={this.state.project.Project_x0020_Document}></ProjectDocument>
                <div className="clearfix"></div>
                <ProjectTeamMembers projectTeamMembers={this.state.project.Project_x0020_Team_x0020_Members}></ProjectTeamMembers>
                <ProjectProjectRoleResponsibility projectRoleResponsibility={this.state.project.Project_x0020_Infromation} ></ProjectProjectRoleResponsibility>
                <ProjectTeamConversation></ProjectTeamConversation>
              </div>
            </div>
          </section>
        </div>
      </div>
    );
  }
}
