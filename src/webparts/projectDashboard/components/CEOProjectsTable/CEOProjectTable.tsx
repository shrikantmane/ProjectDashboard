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
import { ProjectTimeLine, Groups, TimeLineItems }from "../CEOProjectTimeLine/ProjectTimeLine";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import moment from 'moment/src/moment';

export default class CEOProjectTable extends React.Component<
  ICEOProjectProps,
  ICEOProjectState
> {
  constructor(props) {
    super(props);
    this.state = {
      projectList: new Array<CEOProjects>(),
      projectTimeLine: new ProjectTimeLine,
      projectName: null,
      ownerName: null,
      status: null,
      priority: null,
      isLoading : true,
      isTeamMemberLoaded : false,
      isKeyDocumentLoaded :false,
      isTagLoaded :false,
      expandedRowID: -1,
      expandedRows :[]
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
        <b>Active Mildstone:</b>      
          {rowData.MildStone
            ?  rowData.MildStone.Title
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
    if (!this.state.isTeamMemberLoaded && !this.state.isTagLoaded && !this.state.isKeyDocumentLoaded && data.ID == this.state.expandedRowID){
      return <div className={styles.spinnerStyling}><i className="fas fa-spinner"></i></div>
    }
    return (
      <div className={styles.milestoneExpand}>
      <div className={styles.expandIndicator}>
            <i className="fas fa-caret-down"></i>
        </div>
        <div className={styles.milestoneHeader}>
          <div className="row">
            <div className="col-md-2 col-12">
              <div className={styles.milestoneHeader}>
                <div>Milestone List</div>
                {/* <div className={styles.activityStatus}>
                  Last Activity Yesterday
                </div> */}
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
            <div className="col-md-4 col-12">
              <button type="button" className="btn btn-outline btn-sm" style={{backgroundColor : "#1b1a30" , border:"1px solid #504f6c", color: "#fff", fontSize:"12px"}}>
                Project Outline
              </button>
            </div>
          </div>
        </div>
        <table className={styles.milestoneList} style={{ width: "100%" }}>
          <col style={{ width: "25%" }} />
          <col style={{ width: "15%" }} />
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
                       { item.Status0 && item.Status0.Status != "" ? 
                        <div className={styles.statusDetail}>
                          <div
                            className={
                              styles.statusPill + " " + styles.completeStatus
                            }
                            style={{
                              backgroundColor: item.Status0.Status_x0020_Color
                            }}
                          >
                            { item.Status0.Status }
                          </div>
                        </div> :
                        null
                       }
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
                          { item.Team_x0020_Member ? 
                            <img src={item.Team_x0020_Member.ImgUrl} title={item.Team_x0020_Member.Title} />
                            :null
                          }                           
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
                              iconClass = "far fa-file-word";
                              break;
                            case "pdf":
                              iconClass = "far fa-file-pdf";
                              break;
                            case "xls":
                            case "xlsx":
                              iconClass = "far fa-file-excel";
                              break;
                            case "png":
                            case "jpeg":
                            case "gif":
                              iconClass = "far fa-file-image";
                              break;
                            default:
                              iconClass = "fa fa-file";
                              break;
                          }
                          return (
                            <div className={styles.fileName}>
                              <i
                                className={iconClass}
                                style={{ marginRight: "5px" }}
                              />
                              <a href={item.File.ServerRelativeUrl} target="_blank">
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
      // this.getMildStonesByProject(event.data[event.data.length - 1]);
      this.getKeyDocumentsByProject(event.data[event.data.length - 1]);
      this.getTaggingByProject(event.data[event.data.length - 1]);
      this.getTeamMembersByProject(event.data[event.data.length - 1]);      
      if(this.state.expandedRows <  event.data){
        this.setState({ isTeamMemberLoaded :false ,isKeyDocumentLoaded: false, isTagLoaded:false, expandedRowID : event.data[event.data.length - 1].ID });
      }
    }
    this.setState({ expandedRows: event.data}); 
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
        let timeLine = new ProjectTimeLine();
        timeLine = this.getProjectTimeLineData(filterdRecords);

    this.setState({
      globalFilter: event.target.value,
      projectTimeLine: timeLine
    });
  }

  private getProjectTimeLineData (projects:Array<CEOProjects>){
        let timeLine = new ProjectTimeLine();
        let groups = new Array<Groups>();
        let timeLineItems =new Array<TimeLineItems>();
        projects.forEach(element => {
          groups.push({
            id : element.ID,
            title : element.Project
          });
          element.MildStoneList.forEach(mildstone => {
            timeLineItems.push({
              id: mildstone.ID,
              group: element.ID,
              title: mildstone.Title,
              start_time: moment(new Date(mildstone.StartDate).setHours(0,0,0,0)),
              end_time: moment(new Date(mildstone.DueDate).setHours(23,59,59,59)) 
            })
          });         
        });
        timeLine.groups = groups;
        timeLine.items = timeLineItems;
        return timeLine;
  }

  onProjectNameChange(event) {
    this.dt.filter(event.target.value, "Project", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.Project.toLowerCase().match(event.target.value.toLowerCase());
    });    

    let timeLine = new ProjectTimeLine();
        timeLine = this.getProjectTimeLineData(filterdRecords);

    this.setState({
      projectName: event.target.value,
      projectTimeLine: timeLine
    });
  }

  onOwnerNameChange(event) {
    this.dt.filter(event.target.value, "OwnerTitle", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.OwnerTitle.toLowerCase().match(event.target.value.toLowerCase());
    });

    let timeLine = new ProjectTimeLine();
    timeLine = this.getProjectTimeLineData(filterdRecords);

    this.setState({
      ownerName: event.target.value,
      projectTimeLine: timeLine
    });
  }

  onStatusChange(event) {
    this.dt.filter(event.target.value, "StatusText", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.StatusText.toLowerCase().match(event.target.value.toLowerCase());
    });    
    
    let timeLine = new ProjectTimeLine();
        timeLine = this.getProjectTimeLineData(filterdRecords);

    this.setState({
      status: event.target.value,
      projectTimeLine: timeLine
    });
  }
  onPrioritychange(event) {
    this.dt.filter(event.target.value, "Priority", "contains");
    let filterdRecords = this.state.projectList.filter(item => {
      return item.Priority.toLowerCase().match(event.target.value.toLowerCase());
    });
    
    let timeLine = new ProjectTimeLine();
    timeLine = this.getProjectTimeLineData(filterdRecords);

    this.setState({
      priority: event.target.value,
      projectTimeLine: timeLine
    });
  }
  /* Html UI */

  public render(): React.ReactElement<ICEOProjectProps> {
    var header = (
      <div>
      <label className={styles.globalHeading}>CEO Dashboard</label>
      <div className={styles.globalSearch} style={{ textAlign: "left" }}>       
        <input
          type="text"
          placeholder="Search"
          onChange={this.handleGlobalSearchChange}          
        />
         <i className="fa fa-search" style={{ margin: "4px 4px 0 5px" }} />
      </div>
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

    var milstoneFilter = (
      <input
        type="text"
        className={styles.filterCustom}
        style={{ width: "100%", visibility:"hidden", margin:"-5px" }}       
      />
    );

    return (
      <div className={styles.CEOProjectDashboard}>
      {
        !this.state.isLoading ?
       <div>
        {this.state.projectTimeLine &&  this.state.projectTimeLine.groups.length > 0 ? (
          <CEOProjectTimeLine groups={this.state.projectTimeLine.groups} items ={this.state.projectTimeLine.items} />
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
            responsive={true}
            className={styles.datatablePosition}
            expandedRows={this.state.expandedRows}
            onRowToggle={this.onRowToggle.bind(this)}
            rowExpansionTemplate={this.rowExpansionTemplate.bind(this)}
          >
            <Column
              expander={true}
              style={{ width: "3em" }}
              className={styles.firstColExpand}
            />
            <Column
              field="Project"
              header="Project Name"
              body={this.projectNameTemplate}
              style={{ width: "26%" }}
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
              filter={true}
              filterElement={milstoneFilter}
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
        : <div style={{textAlign : "center", fontSize:"25px"}}><i className="fas fa-spinner"></i></div>
        }
      </div>
    );
  }

  /* Api Call*/

  private getProjectList(): void {
    sp.web.lists
      .getByTitle("Project")
      .items.select(
        "ID",
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
      .items.select(
        "ID",
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
        let timeline = new ProjectTimeLine() ;
        let groups =  Array<Groups>();
        let timeLineItems = Array<TimeLineItems>();
        projectList.forEach(item => {

          groups.push({
            id : item.ID,
            title : item.Project
          });

          let filteredMilestones = filter(milestones, function(milstoneItem) {
            return milstoneItem.Project && milstoneItem.Project.ID == item.ID;
          });
          let mildstone = null;
          let mildstones = [];
          let currentDate = new Date(new Date().setHours(0,0,0,0));
          let lastDueDate = new Date(new Date().setHours(0,0,0,0));

          filteredMilestones = sortBy(filteredMilestones, function(dateObj) {
            return new Date(dateObj.StartDate);
          });
          for (let count = 0; count < filteredMilestones.length; count++) {

            timeLineItems.push({
              id: filteredMilestones[count].ID,
              group: item.ID,
              title: filteredMilestones[count].Title,
              start_time: moment(new Date(filteredMilestones[count].StartDate).setHours(0,0,0,0)),
              end_time: moment(new Date(filteredMilestones[count].DueDate).setHours(23,59,59,59)) 
            })

            let mStartDate = new Date(new Date(filteredMilestones[count].StartDate).setHours(0,0,0,0));
            let mDueDate = new Date(new Date(filteredMilestones[count].DueDate).setHours(0,0,0,0));
            
            if (currentDate >= mStartDate && currentDate <= mDueDate) {
              mildstone = filteredMilestones[count];
            } else {
              if (currentDate < mStartDate) {
                mildstones.push(filteredMilestones[count]);
              }
            }

            if(filteredMilestones[count].AssignedTo && filteredMilestones[count].AssignedTo.length > 0){
              filteredMilestones[count].AssignedTo.forEach(element => {
                if (element.EMail != null) {
                  element.imgURL =
                        "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
                        element.EMail +
                        "&UA=0&size=HR64x64&sc=1531997060853";
                   // element.imgURL = "/_layouts/15/userphoto.aspx?size=S&username=" + element.EMail;       
                } else {
                 element.imgURL = "";               
                }
              });
            }
          }
          item.MildStoneList = filteredMilestones;
          item.MildStone =
            mildstone == null && mildstones.length > 0
              ? mildstones[0]
              : mildstone;       
            if (item.AssignedTo && item.AssignedTo.length > 0){
                  item.AssignedTo.forEach(element => {
                    if (element.EMail != null) {
                      element.imgURL =
                      "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
                      element.EMail +
                      "&UA=0&size=HR64x64&sc=1531997060853";                     
                    } else {
                      element.imgURL = "";
                    }

                    // if (element.EMail != null) {
                    //   element.imgURL = "/_layouts/15/userphoto.aspx?size=S&username=" + element.EMail                     
                    // } else {
                    //   element.imgURL = "";
                    // }
                  });
            }

          item.StatusText = item.Status0 ? item.Status0.Status : "";
          item.OwnerTitle =
            item.AssignedTo && item.AssignedTo.length > 0
              ? item.AssignedTo[0].Title
              : "";

        });      
        timeline.groups = groups;
        timeline.items = timeLineItems;
        this.setState({ projectList: projectList, projectTimeLine: timeline, isLoading:false });
      });
  }

  private getTeamMembersByProject(currentProject :CEOProjects): void {
  let filter = "Project/ID eq " + currentProject.ID + " and  Status eq 'Active'";
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
            if (item.Team_x0020_Member.EMail) {             
                item.Team_x0020_Member.ImgUrl =
                "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
                item.Team_x0020_Member.EMail +
                "&UA=0&size=HR64x64&sc=1531997060853";
               // item.Team_x0020_Member.ImgUrl = "/_layouts/15/userphoto.aspx?size=S&username=" + item.Team_x0020_Member.EMail     
            } else {
              //item.Team_x0020_Member.ImgUrl = "";
              item.Team_x0020_Member.ImgUrl = "";
                       
            }
          }
        });
        let projects = this.state.projectList;
        let project = find(projects, { ID: currentProject.ID });
        project.TeamMemberList = response;
        this.setState({ projectList: projects,  isTeamMemberLoaded : true  });
      });
  }
  private getMildStonesByProject(currentProject :CEOProjects): void {
    let filter = "Project/ID eq '" + currentProject.ID + "' and Duration eq 0";
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
          if(item.AssignedTo && item.AssignedTo.length > 0){
            item.AssignedTo.forEach(element => {
              if (element.EMail != null) {
                element.imgURL =
                      "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
                      element.EMail +
                      "&UA=0&size=HR64x64&sc=1531997060853";
               // element.imgURL ="/_layouts/15/userphoto.aspx?size=S&username=" + element.EMail  
              } else {
                element.imgURL ="";
              }
            });
          }
        });
        let projects = this.state.projectList;
        let project = find(projects, { ID: currentProject.ID });
        project.MildStoneList = response;
        this.setState({ projectList: projects });
      });
  }

  private getKeyDocumentsByProject(currentProject :CEOProjects): void {
    let filter = "Project/ID eq '" + currentProject.ID + "'";
    sp.web.lists
      .getByTitle("Project Documents")
      .items.select("File", "Project/ID", "Project/Title")
      .expand("File", "Project")
      .filter(filter)
      .get()
      .then((response: Array<Documents>) => {
        let projects = this.state.projectList;
        let project = find(projects, { ID: currentProject.ID });
        project.DocumentList = response;
        this.setState({ projectList: projects, isKeyDocumentLoaded : true });
      });
  }

  private getTaggingByProject(currentProject :CEOProjects): void {
    let filter = "Project/ID eq '" + currentProject.ID + "'";
    sp.web.lists
      .getByTitle("Project Tags")
      .items.select("ID", "Tags", "Color")
      .filter(filter)
      .get()
      .then((response: Array<Tags>) => {
        if (response != null && response.length > 0) {
          let projects = this.state.projectList;
          let project = find(projects, { ID: currentProject.ID });
          project.TagList = response;
          this.setState({ projectList: projects, isTagLoaded : true  });
        }
      });
  }
}
