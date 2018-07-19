import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectDashboard.module.scss";
import { ICEOProjectProps } from "./ICEOProjectProps";
import { ICEOProjectState } from "./ICEOProjectState";
import {
  CEOProjects,
  MildStones,
  TeamMembers,
  Tags,
  Documents
} from "./CEOProject";
import CEOProjectTimeLine from "../CEOProjectTimeLine/CEOProjectTimeLine";
import ProjectTimeLine from "../CEOProjectTimeLine/ProjectTimeLine";
import { find, filter } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";

export default class CEOProjectTable extends React.Component<
  ICEOProjectProps,
  ICEOProjectState
> {
  constructor(props) {
    super(props);
    this.state = {
      projectList: new Array<CEOProjects>(),
      projectTimeLine: new Array<ProjectTimeLine>(),
      projectName: null,
      ownerName: null,
      status: null,
      priority: null
    };
    this.handleGlobalSearchChange = this.handleGlobalSearchChange.bind(this);
    this.onProjectNameChange = this.onProjectNameChange.bind(this);
    this.onOwnerNameChange = this.onOwnerNameChange.bind(this);
    this.onStatusChange = this.onStatusChange.bind(this);
    this.onPrioritychange = this.onPrioritychange.bind(this);
  }
  dt: any;
  componentDidMount() {
    SPComponentLoader.loadCss(
      "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
    );
    this.getProjectList();
  }
  componentWillReceiveProps(nextProps) {}

  /* Private Methods */

  private projectNameTemplate(rowData: CEOProjects, column) {
    return (
      <div>
        <div className={styles.projectName}>{rowData.Project}</div>
        <div>
          {" "}
          {rowData.MildStone
            ? "Active Mildstone: " + rowData.MildStone.Title
            : ""}
        </div>
      </div>
    );
  }

  private ownerTemplate(rowData: CEOProjects, column) {
    if (rowData.AssignedTo)
      return (
        <div className={styles.ownerImg}>
          <img src={rowData.AssignedTo[0].imgURL} />
          <div>{rowData.AssignedTo[0].Title}</div>
        </div>
      );
  }

  private mildstoneTemplate(rowData: CEOProjects, column) {
    let startDate: any;
    let dueDate: any;
    if (rowData.MildStone != null) {
      startDate = rowData.MildStone.StartDate
        ? new Date(rowData.MildStone.StartDate).toDateString()
        : "";
      dueDate = rowData.MildStone.DueDate
        ? new Date(rowData.MildStone.DueDate).toDateString()
        : "";
      return (
        <div className="row">
          <div className="col-md-12 col-12">
            <div className={styles.milestoneDetail}>
              <div className={styles.innerCol}>
                <div className={styles.milestonedate}>
                  {startDate != "" && dueDate != ""
                    ? startDate + "-" + dueDate
                    : ""}
                </div>
                <div
                  className={styles.blueMilestone + " " + styles.stepMilestone}
                >
                  <label title={rowData.MildStone.Body}>
                    {rowData.MildStone.Body}
                  </label>
                  <span className={styles.milestoneStatus}>
                    {rowData.MildStone.Status0
                      ? rowData.MildStone.Status0.Status
                      : ""}
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
    }
  }

  private statusTemplate(rowData: CEOProjects, column) {
    if (rowData.Status0)
      return (
        <div className={styles.statusDetail}>
          <div
            className={styles.completeStatus + " " + styles.statusPill}
            style={{ backgroundColor: rowData.Status0.Status_x0020_Color }}
          >
            {rowData.Status0.Status}
          </div>
        </div>
      );
  }

  private priorityTemplate(rowData: CEOProjects, column) {
    return <div className={styles.priorityDetail}>{rowData.Priority}</div>;
  }

  private rowExpansionTemplate(data: CEOProjects) {
    return (
      <div className={styles.milestoneExpand}>
        <div className={styles.milestoneHeader}>
          <div className="row">
            <div className="col-md-3 col-12">
              <div className={styles.milestoneHeader}>
                <div>Milestone List</div>
                <div className={styles.activityStatus}>
                  Last Activity Yesterday
                </div>
              </div>
            </div>
            <div className="col-md-6 col-12">
              <div className={styles.tagList}>
                {data.TagList != null
                  ? data.TagList.map((item, key) => {
                      return (
                        <span className={styles.pinkTag} style={{backgroundColor : item.Color}}>{item.Tags}</span>
                      );
                    })
                  : null}
              </div>
            </div>
            <div className="col-md-3 col-12">
              <button type="button" className="btn btn-outline btn-sm">
                Project Outline
              </button>
            </div>
          </div>
        </div>
        <table className={styles.milestoneList} style={{ width: "100%" }}>
          <col style={{ width: "30%" }} />
          <col style={{ width: "10%" }} />
          <col style={{ width: "35%" }} />
          <col style={{ width: "15%" }} />
          <col style={{ width: "10%" }} />
          <thead>
            <tr>
              <th>Milestone Name</th>
              <th>Owner</th>
              <th>Description</th>
              <th>Start Date</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {data.MildStoneList != null
              ? data.MildStoneList.map((item, key) => {
                  return (
                    <tr className={styles.milestoneItems}>
                      <td>
                        <div className="milestoneName">{item.Title}</div>
                      </td>
                      <td>
                        <div className={styles.ownerImg}>
                          <img
                            src={
                              item.AssignedTo && item.AssignedTo.length > 0
                                ? item.AssignedTo[0].imgURL
                                : ""
                            }
                            className="img-responsive"
                          />
                          <div>
                            {item.AssignedTo && item.AssignedTo.length > 0
                              ? item.AssignedTo[0].Title
                              : ""}
                          </div>
                        </div>
                      </td>
                      <td>
                        <div className={styles.milestoneDesc}>
                          <label title={item.Body}>{item.Body}</label>
                        </div>
                      </td>
                      <td>
                        <div>
                          {item.StartDate
                            ? new Date(item.StartDate).toDateString()
                            : ""}
                        </div>
                      </td>
                      {/* <td>
                              <div>{item.Status0 ? item.Status0.Status : ""}</div>
                          </td> */}
                      <td>
                        <div className={styles.statusDetail}>
                          <div
                            className={
                              styles.statusPill + " " + styles.completeStatus
                            }
                            style={{
                              backgroundColor: item.Status0.Status_x0020_Color
                            }}
                          >
                            {item.Status0 ? item.Status0.Status : ""}
                          </div>
                        </div>
                      </td>
                    </tr>
                  );
                })
              : null}
          </tbody>
        </table>
        <div className={styles.milestoneFooter}>
          <div className="row">
            <div className="col-md-4 col-12">
              <div className={styles.teamMembers}>
                <h5>Team Members</h5>
                <div className={styles.memberList}>
                  {data.TeamMemberList != null
                    ? data.TeamMemberList.map((item, key) => {
                        return (
                          <div className={styles.memberImg}>
                            <img src={item.Team_x0020_Member.ImgUrl} />
                            {/* <span className={styles.badgeLight}>17</span> */}
                          </div>
                        );
                      })
                    : null}
                </div>
              </div>
            </div>
            <div className="col-md-5 col-12">
              <div className={styles.keyDoc}>
                <h5>Key Documents</h5>
                <div className={styles.docList}>
                  {data.DocumentList != null
                    ? data.DocumentList.map((item, key) => {
                        if (item.File) {
                          let type = "";
                          let iconClass = "";
                          let data = item.File.Name.split(".");
                          if (data.length > 1) {
                            type = data[1];
                          }
                          switch (type.toLowerCase()) {
                            case "doc":
                            case "docx":
                              iconClass = "far fa-file-excel";
                              break;
                            case "pdf":
                              iconClass = "far fa-file-pdf";
                              break;
                            case "xls":
                              iconClass = "far fa-file-excel";
                              break;
                            case "img":
                              iconClass = "far fa-file-image";
                              break;
                            default:
                              iconClass = "";
                              break;
                          }
                          return (
                            <div className={styles.fileName}>
                              <i
                                className={iconClass}
                                style={{ marginRight: "5px" }}
                              />
                              <a href={item.File.LinkingUri} target="_blank">
                                {item.File.Name}
                              </a>
                            </div>
                          );
                        }
                      })
                    : null}
                </div>
              </div>
            </div>
            <div className="col-md-3 col-12">
              <div className={styles.projectPageLink}>
                <h5>For Detailed Overview Go To</h5>
                <button type="button" className="btn btn-white btn-sm">
                  Project Page
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private onRowToggle(event) {
    if (event.data && event.data.length > 0) {
      this.getMildStonesByProject(event.data[event.data.length - 1].Project);
      this.getKeyDocumentsByProject(event.data[event.data.length - 1].Project);
      this.getTaggingByProject(event.data[event.data.length - 1].Project);
      this.getTeamMembersByProject(event.data[event.data.length - 1].Project);
    }
    this.setState({ expandedRows: event.data });
  }

  private handleGlobalSearchChange(event) {
    let filterdRecords = this.state.projectList.filter(item => {
      return (
        item.Project.toLowerCase().match(event.target.value.toLowerCase()) ||
        item.OwnerTitle.toLowerCase().match(event.target.value.toLowerCase()) ||
        item.Priority.toLowerCase().match(event.target.value.toLowerCase()) ||
        item.StatusText.toLowerCase().match(event.target.value.toLowerCase())
      );
    });
    let timeLines = [];
    filterdRecords.forEach(project => {
      timeLines.push({
        id: project.Project_x0020_ID,
        name: project.Project,
        start: project.StartDate,
        end: project.DueDate,
      });
    });

    this.setState({
      globalFilter: event.target.value,
      projectTimeLine: timeLines
    });
  }

  onProjectNameChange(event) {
    this.dt.filter(event.target.value, "Project", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.Project.toLowerCase().match(event.target.value.toLowerCase());
    });

    let timeLines = new Array<ProjectTimeLine>();
    filterdRecords.forEach(project => {
      timeLines.push({
        id: project.Project_x0020_ID,
        name: project.Project,
        start: project.StartDate,
        end: project.DueDate,
      });
    });
    this.setState({
      projectName: event.target.value,
      projectTimeLine: timeLines
    });
  }

  onOwnerNameChange(event) {
    this.dt.filter(event.target.value, "OwnerTitle", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.OwnerTitle.toLowerCase().match(event.target.value.toLowerCase());
    });

    let timeLines = new Array<ProjectTimeLine>();
    filterdRecords.forEach(project => {
      timeLines.push({
        id: project.Project_x0020_ID,
        name: project.Project,
        start: project.StartDate,
        end: project.DueDate,
      });
    });
    this.setState({
      ownerName: event.target.value,
      projectTimeLine: timeLines
    });
  }

  onStatusChange(event) {
    this.dt.filter(event.target.value, "StatusText", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.StatusText.toLowerCase().match(event.target.value.toLowerCase());
    });

    let timeLines = new Array<ProjectTimeLine>();
    filterdRecords.forEach(project => {
      timeLines.push({
        id: project.Project_x0020_ID,
        name: project.Project,
        start: project.StartDate,
        end: project.DueDate,
      });
    });
    this.setState({ status: event.target.value, projectTimeLine: timeLines });
  }
  onPrioritychange(event) {
    this.dt.filter(event.target.value, "Priority", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.Priority.toLowerCase().match(event.target.value.toLowerCase());
    });

    let timeLines = new Array<ProjectTimeLine>();
    filterdRecords.forEach(project => {
      timeLines.push({
        id: project.Project_x0020_ID,
        name: project.Project,
        start: project.StartDate,
        end: project.DueDate,
      });
    });
    this.setState({ priority: event.target.value, projectTimeLine: timeLines });
  }
  /* Html UI */

  public render(): React.ReactElement<ICEOProjectProps> {
    var header = (
      <div className={styles.globalSearch} style={{ textAlign: "left" }}>       
        <input
          type="text"
          placeholder="Search"
          onChange={this.handleGlobalSearchChange}          
        />
         <i className="fa fa-search" style={{ margin: "4px 4px 0 5px" }} />
      </div>
    );

    // let brandFilter = <input className={styles.filterCustom}/>

    var projectNameFilter = (
      <input
        type="text"
        style={{ width: "100%" }}
        className={styles.filterCustom}
        value={this.state.projectName}
        onChange={this.onProjectNameChange}
      />
    );

    var ownerNameFilter = (
      <input
        type="text"
        style={{ width: "100%" }}
        className={styles.filterCustom}
        value={this.state.ownerName}
        onChange={this.onOwnerNameChange}
      />
    );

    var statusFilter = (
      <input
        type="text"
        style={{ width: "100%" }}
        className={styles.filterCustom}
        value={this.state.status}
        onChange={this.onStatusChange}
      />
    );

    var priorityFilter = (
      <input
        type="text"
        style={{ width: "100%" }}
        className={styles.filterCustom}
        value={this.state.priority}
        onChange={this.onPrioritychange}
      />
    );

    return (
      <div className={styles.CEOProjectDashboard}>
        {this.state.projectTimeLine.length > 0 ? (
          <CEOProjectTimeLine tasks={this.state.projectTimeLine} />
        ) : null}
        <div style={{ marginTop: "10px" }}>
          <DataTable
            paginator={true}
            rows={5}
            rowsPerPageOptions={[5, 10, 20]}
            ref={el => (this.dt = el)}
            globalFilter={this.state.globalFilter}
            header={header}
            value={this.state.projectList}
            // responsive={true}
            className={styles.datatablePosition}
            expandedRows={this.state.expandedRows}
            onRowToggle={this.onRowToggle.bind(this)}
            rowExpansionTemplate={this.rowExpansionTemplate.bind(this)}
          >
            <Column
              expander={true}
              style={{ width: "2em" }}
              className={styles.firstColExpand}
            />
            <Column
              field="Project"
              header="Project Name"
              body={this.projectNameTemplate}
              style={{ width: "30%" }}
              filter={true}
              sortable={true}
              filterElement={projectNameFilter}
            />
            <Column
              field="OwnerTitle"
              header="Owner"
              body={this.ownerTemplate}
              style={{ width: "20%" }}
              filter={true}
              sortable={true}
              filterElement={ownerNameFilter}
            />
            <Column
              field="MildStone"
              header="Mildstone"
              body={this.mildstoneTemplate}
              style={{ width: "30%" }}
            />
            <Column
              field="StatusText"
              header="Status"
              body={this.statusTemplate}
              style={{ width: "10%" }}
              filter={true}
              sortable={true}
              filterElement={statusFilter}
            />
            <Column
              field="Priority"
              header="Priority"
              body={this.priorityTemplate}
              style={{ width: "10%" }}
              filter={true}
              sortable={true}
              filterElement={priorityFilter}
            />
          </DataTable>
        </div>
      </div>
    );
  }

  /* Api Call*/

  private getProjectList(): void {
    sp.web.lists
      .getByTitle("Project")
      .items.select(
        "Project_x0020_ID",
        "Project",
        "StartDate",
        "DueDate",
        "AssignedTo/Title",
        "AssignedTo/ID",
        "AssignedTo/EMail",
        "Status0/ID",
        "Status0/Status",
        "Status0/Status_x0020_Color",
        "Priority",
        "Body"
      )
      .expand("AssignedTo", "Status0")
      .getAll()
      .then((response: Array<CEOProjects>) => {
        this.getMildStones(response);
      });
  }

  private getMildStones(projectList: Array<CEOProjects>): void {
    sp.web.lists
      .getByTitle("Tasks List")
      // .items
      .items.select(
        "Title",
        "StartDate",
        "DueDate",
        "Status0/ID",
        "Body",
        "Status0/Status",
        "Status0/Status_x0020_Color",
        "Project/ID",
        "Project/Title",
        "AssignedTo/Title",
        "AssignedTo/ID",
        "AssignedTo/EMail"
      )
      .expand("Project", "Status0", "AssignedTo")
      .filter("Duration eq 0")
      .get()
      .then((milestones: Array<MildStones>) => {
        let timeline = new Array<ProjectTimeLine>();
        projectList.forEach(item => {
          let filteredMilestones = filter(milestones, function(milstoneItem) {
            return milstoneItem.Project.ID.toString() == item.Project_x0020_ID;
          });
          let mildstone = null;
          let mildstones = [];
          let currentDate = new Date().getTime();
          for (let count = 0; count < filteredMilestones.length; count++) {
            let startDate = new Date(
              filteredMilestones[count].StartDate
            ).getTime();
            let dueDate = new Date(filteredMilestones[count].DueDate).getTime();

            if (currentDate <= startDate && currentDate >= dueDate) {
              mildstone = filteredMilestones[count];
            } else {
              if (startDate > currentDate) {
                mildstones.push(filteredMilestones[count]);
              }
            }
          }
          item.MildStone =
            mildstone == null && mildstones.length > 0
              ? mildstones[0]
              : mildstone;
          item.AssignedTo.forEach(element => {
            if (element.EMail != null) {
              element.imgURL =
                "https://esplrms-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/" +
                element.EMail.split("@")[0].toLowerCase() +
                "_esplrms_onmicrosoft_com_MThumb.jpg";
            } else {
              element.imgURL =
                "https://esplrms.sharepoint.com/sites/projects/SiteAssets/default.jpg";
            }
          });

          item.StatusText = item.Status0 ? item.Status0.Status : "";
          item.OwnerTitle =
            item.AssignedTo && item.AssignedTo.length > 0
              ? item.AssignedTo[0].Title
              : "";

          // Time Line
          timeline.push({
            id: item.Project_x0020_ID,
            name: item.Project,
            start: item.StartDate,
            end: item.DueDate,
          });
        });
        this.setState({ projectList: projectList, projectTimeLine: timeline });
      });
  }

  private getTeamMembersByProject(name): void {
    let filter = "Project/Title eq '" + name + "' and  Status eq 'Active'";
    sp.web.lists
      .getByTitle("Project Team Members")
      .items.select(
        "Team_x0020_Member/ID",
        "Team_x0020_Member/Title",
        "Team_x0020_Member/EMail",
        "Start_x0020_Date",
        "End_x0020_Date",
        "Status",
        "Project/ID",
        "Project/Title"
      )
      .expand("Team_x0020_Member", "Project")
      .filter(filter)
      .get()
      .then((response: Array<TeamMembers>) => {
        response.forEach(item => {
          if (item.Team_x0020_Member) {
            if (item.Team_x0020_Member.Email) {
              item.Team_x0020_Member.ImgUrl =
                "https://esplrms-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/" +
                item.Team_x0020_Member.Email.split("@")[0].toLowerCase() +
                "_esplrms_onmicrosoft_com_MThumb.jpg";
            } else {
              item.Team_x0020_Member.ImgUrl =
                "https://esplrms.sharepoint.com/sites/projects/SiteAssets/default.jpg";
            }
          }
        });
        let projects = this.state.projectList;
        let project = find(projects, { Project: name });
        project.TeamMemberList = response;
        this.setState({ projectList: projects });
      });
  }
  private getMildStonesByProject(name): void {
    let filter = "Project/Title eq '" + name + "' and Duration eq 0";
    sp.web.lists
      .getByTitle("Tasks List")
      .items.select(
        "Title",
        "StartDate",
        "DueDate",
        "Body",
        "Status0/ID",
        "Status0/Status",
        "Status0/Status_x0020_Color",
        "Project/ID",
        "Project/Title",
        "AssignedTo/Title",
        "AssignedTo/ID",
        "AssignedTo/EMail",
        "Priority"
      )
      .expand("Project", "Status0", "AssignedTo")
      .filter(filter)
      .get()
      .then((response: Array<MildStones>) => {
        response.forEach(item => {
          item.AssignedTo.forEach(element => {
            if (element.EMail != null) {
              element.imgURL =
                "https://esplrms-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/" +
                element.EMail.split("@")[0].toLowerCase() +
                "_esplrms_onmicrosoft_com_MThumb.jpg";
            } else {
              element.imgURL =
                "https://esplrms.sharepoint.com/sites/projects/SiteAssets/default.jpg";
            }
          });
        });
        let projects = this.state.projectList;
        let project = find(projects, { Project: name });
        project.MildStoneList = response;
        this.setState({ projectList: projects });
      });
  }

  private getKeyDocumentsByProject(name): void {
    let filter = "Project/Title eq '" + name + "'";
    sp.web.lists
      .getByTitle("Project Documents")
      .items.select("File", "Project/ID", "Project/Title")
      .expand("File", "Project")
      .filter(filter)
      .get()
      .then((response: Array<Documents>) => {
        let projects = this.state.projectList;
        let project = find(projects, { Project: name });
        project.DocumentList = response;
        this.setState({ projectList: projects });
      });
  }

  private getTaggingByProject(name): void {
    let filter = "Project/Title eq '" + name + "'";
    sp.web.lists
      .getByTitle("Project Tags")
      .items.select("ID", "Tags", "Color")
      .filter(filter)
      .get()
      .then((response: Array<Tags>) => {
        if (response != null && response.length > 0) {
          let projects = this.state.projectList;
          let project = find(projects, { Project: name });
          project.TagList = response;
          this.setState({ projectList: projects });
        }
      });
  }
}
