import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from '../ProjectDashboard.module.scss';
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
    };
  }

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
          <div className= {styles.projectName}>{rowData.Project}</div>
          <div> {rowData.MildStone ? "Active Mildstone: " + rowData.MildStone.Title : ""}</div>
        </div>
      );
  }

  private ownerTemplate(rowData: CEOProjects, column) {
    if (rowData.AssignedTo)
      return (
        // <div>
        //   <img
        //     src={rowData.AssignedTo[0].imgURL}
        //     style={{ marginRight: "5px", width: "20px" }}
        //   />
        //   {rowData.AssignedTo[0].Title}
        // </div>
        <div className={styles.ownerImg}>
         <img src={rowData.AssignedTo[0].imgURL} />
         <div>{rowData.AssignedTo[0].Title}</div>
        </div>
      );
  }

private mildstoneTemplate(rowData: CEOProjects, column) {
  let startDate : any ;
  let dueDate : any ;
  if(rowData.MildStone != null )  {
  startDate = rowData.MildStone.StartDate ? new Date(rowData.MildStone.StartDate).toDateString() : ""; 
  dueDate = rowData.MildStone.DueDate ? new Date(rowData.MildStone.DueDate).toDateString() : ""; 
  return (
    <div className="row">
    <div className="col-md-12 col-12">
        <div className={styles.milestoneDetail}>
            <div className={styles.innerCol}>
                <div className={styles.milestonedate}>{startDate != "" && dueDate != "" ? startDate + "-" + dueDate : ""}</div>
                <div className={styles.blueMilestone + ' ' + styles.stepMilestone}> 
                    <label>{rowData.MildStone.Body}</label>
                    <span className={styles.milestoneStatus}>{rowData.MildStone.Status0? rowData.MildStone.Status0.Status :""}</span>
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
      // <div
      //   style={{
      //     backgroundColor: rowData.Status0.Status_x0020_Color,
      //     height: "2.9em",
      //     width: "100%",
      //     textAlign: "center",
      //     paddingTop: 7,
      //     color: "#fff"
      //   }}
      // >
      //   {rowData.Status0.Status}
      // </div>
        <div className={styles.statusDetail}>
          <div className= {styles.completeStatus +' ' + styles.statusPill} style={{backgroundColor: rowData.Status0.Status_x0020_Color}}>{rowData.Status0.Status}</div>
      </div>
    );
}

private priorityTemplate (rowData: CEOProjects, column){
return (<div className={styles.priorityDetail}>{rowData.Priority}</div>);
}

  private rowExpansionTemplate(data : CEOProjects) {
    let teamMemberList = new Array<TeamMembers>();
    let extryTeamMemberList = new Array<TeamMembers>();
    let tagList = new Array<Tags>();
    let extryTagList = new Array<Tags>();
    let docList = new Array<Documents>();
    let extryDocList = new Array<Documents>();
    // Team Members
    if(data.TeamMemberList != null){
      if(data.TeamMemberList.length > 4){
        data.TeamMemberList.forEach((element, index) => {
          if(index < 4){
          teamMemberList.push(element);
          }else{
            extryTeamMemberList.push(element);
          }
        });
      }else {
        teamMemberList = data.TeamMemberList;
      }
    }
   
    // Tags
    if(data.TagList != null){
      if(data.TagList.length > 4){
        data.TagList.forEach((element, index) => {
          if(index < 4){
            tagList.push(element);
          }else{
            extryTagList.push(element);
          }
        });
      }else {
        tagList = data.TagList;
      }
    }

    if(data.DocumentList != null){
      if(data.DocumentList.length > 4){
        data.DocumentList.forEach((element, index) => {
          if(index < 3){
            docList.push(element);
          }else{
            extryDocList.push(element);
          }
        });
      }else {
        docList = data.DocumentList;
      }
    }

    return (
    <div className={styles.milestoneExpand}>
      <div className={styles.milestoneHeader}>
          <div className="row">
              <div className="col-md-3 col-12">
                  <div className={styles.milestoneHeader}>
                      <div>Milestone List</div>
                      <div className={styles.activityStatus}>Last Activity Yesterday</div>
                  </div>
              </div>
              <div className="col-md-6 col-12">
                 <div className={styles.tagList}>
                    {
                      tagList != null ?
                      tagList.map((item, key)=>{
                          return (
                            <span className={styles.pinkTag}>{item.Tags}</span>                                   
                          );
                        })
                        :null
                      }
                    {
                      extryTagList != null && extryTagList.length > 0 ?
                        (<div className={styles.memberImg}>
                            <span className={styles.moreMember}>+{extryTagList.length}</span>
                        </div>)
                        :null
                    }
                  </div>
              </div>
              <div className="col-md-3 col-12">
                  <button type="button" className="btn btn-outline btn-sm">Project Outline</button>
              </div>
          </div>
      </div>
      <table className={styles.milestoneList} style={{ width : "100%"}}>
          <col style={{ width : "30%"}}/>
          <col style={{ width : "10%"}}/>
          <col style={{ width : "35%"}}/>
          <col style={{ width : "15%"}}/>
          <col style={{ width : "10%"}}/>
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
               {
                 data.MildStoneList != null ?
                  data.MildStoneList.map((item, key)=>{
                    return (
                        <tr className={styles.milestoneItems}>
                          <td>
                              <div className="milestoneName">{item.Title}</div>
                          </td>
                          <td>
                              <div className={styles.ownerImg}>
                                  <img src={item.AssignedTo && item.AssignedTo.length > 0 ? item.AssignedTo[0].imgURL : ""} className="img-responsive"/>
                                  <div>{item.AssignedTo && item.AssignedTo.length > 0 ? item.AssignedTo[0].Title : ""}</div>
                              </div>
                          </td>
                          <td>
                              <div className={styles.milestoneDesc}>
                                    {item.Body}
                              </div>
                          </td>
                          <td>
                              <div>{item.StartDate ? new Date(item.StartDate).toDateString() : ""}</div>
                          </td>
                          {/* <td>
                              <div>{item.Status0 ? item.Status0.Status : ""}</div>
                          </td> */}
                          <td>
                              <div className={styles.statusDetail}>
                                  <div className={styles.statusPill + ' ' + styles.completeStatus} style={{backgroundColor: item.Status0.Status_x0020_Color}}>{item.Status0 ? item.Status0.Status : ""}</div>
                              </div>
                          </td>
                      </tr>
                    );
                  })
                  :null
               }              
          </tbody>
      </table>
      <div className={styles.milestoneFooter}>
          <div className="row">
              <div className="col-md-4 col-12">
                  <div className={styles.teamMembers}>
                      <h5>Team Members</h5>
                      <div className={styles.memberList}>
                          {
                            teamMemberList != null ?
                             teamMemberList.map((item, key)=>{
                                return (
                                <div className={styles.memberImg}>
                                    <img src={item.Team_x0020_Member.ImgUrl} />
                                    {/* <span className={styles.badgeLight}>17</span> */}
                                </div>                                    
                                );
                              })
                              :null
                          }
                          {
                            extryTeamMemberList != null && extryTeamMemberList.length > 0 ?
                              (
                                  <span className={styles.moreMember}>+{extryTeamMemberList.length}</span>
                              )
                              :null
                          }
                      </div>
                  </div>
              </div>
              <div className="col-md-5 col-12">
                  <div className={styles.keyDoc}>
                      <h5>Key Documents</h5>
                      <div className={styles.docList}>
                       {
                            docList != null ?
                            docList.map((item, key)=>{
                                if(item.File){
                                  let type = "";
                                  let iconClass ="";
                                  let data = item.File.Name.split(".");
                                  if(data.length > 1){
                                    type = data[1];
                                  }                                  
                                  switch(type.toLowerCase()){
                                    case 'doc':
                                    case 'docx':
                                        iconClass = "far fa-file-pdf";
                                    break;
                                    case 'pdf':
                                        iconClass = "far fa-file-pdf";                                    
                                    break;
                                    case 'xls':
                                        iconClass = "far fa-file-excel";
                                    break
                                    default :
                                      iconClass = ""
                                    break;
                                  }
                                  return (
                                  <div className={styles.fileName}>
                                      <i className={iconClass} style={{marginRight: "5px"}}></i>
                                      <a href={item.File.LinkingUri} target="_blank">{item.File.Name}</a>
                                  </div>                                    
                                  );
                                }
                              })
                              :null
                          }
                           {
                            extryDocList != null && extryDocList.length > 0 ?
                              (
                                  <span className={styles.moreMember}>+{extryDocList.length}</span>
                              )
                              :null
                          }
                      </div>
                  </div>
              </div>
              <div className="col-md-3 col-12">
                  <div className={styles.projectPageLink}>
                      <h5>For Detailed Overview Go To</h5>
                      <button type="button" className="btn btn-white btn-sm">Project Page</button>
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
  /* Html UI */

  
  public render(): React.ReactElement<ICEOProjectProps> {
    var header = <div className= { styles.globalSearch } style={{'textAlign':'left'}}>
                <i className="fa fa-search" style={{margin:'4px 4px 0 0'}}></i>
                <input  type="text" placeholder="Global Search" onChange={(e) => this.setState({globalFilter: e.target.value})}/>
            </div>;

    return (
      <div className={ styles.CEOProjectDashboard }>
        {this.state.projectTimeLine.length > 0 ? (
          <CEOProjectTimeLine tasks={this.state.projectTimeLine} />
        ) : null}
        <div style={{ marginTop: "10px" }}>
          <DataTable paginator={true} rows={5} rowsPerPageOptions={[5,10,20]}
            globalFilter={this.state.globalFilter}
            header={header}
            value={this.state.projectList}
            // responsive={true}
            className={styles.datatablePosition}
            expandedRows={this.state.expandedRows}
            onRowToggle={this.onRowToggle.bind(this)}
            rowExpansionTemplate={this.rowExpansionTemplate.bind(this)}
          >
            <Column expander={true} style={{ width: "2em" }} className={styles.firstColExpand}/>
            <Column field="Project" header="Project Name"  body={this.projectNameTemplate} style={{ width: "30%" }} filter={true} sortable={true}/>            
            <Column field="OwnerTitle" header="Owner" body={this.ownerTemplate} style={{ width: "20%" }} filter={true} sortable={true}/>
            <Column field="MildStone" header="Mildstone" body={this.mildstoneTemplate} style={{ width: "30%" }}/>
            <Column
              field="StatusText"
              header="Status"
              body={this.statusTemplate}
              style={{width:"10%" }}
              filter={true}
              sortable={true}
            />
            <Column field="Priority" header="Priority"  body={this.priorityTemplate} style={{ width: "10%" }} filter={true} sortable={true}/>   
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
      .items.select("Title","StartDate", "DueDate", "Status0/ID", "Body", "Status0/Status", "Status0/Status_x0020_Color","Project/ID", "Project/Title", "AssignedTo/Title", "AssignedTo/ID",
      "AssignedTo/EMail")
      .expand("Project", "Status0","AssignedTo")
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
            let startDate = new Date(filteredMilestones[count].StartDate).getTime();
            let dueDate = new Date(filteredMilestones[count].DueDate).getTime();                         
             
            if(currentDate <= startDate && currentDate >= dueDate ){
              mildstone = filteredMilestones[count];
            }else {
              if(startDate > currentDate ){
                mildstones.push(filteredMilestones[count]);
              }
            }
          }
          item.MildStone = mildstone == null && mildstones.length > 0 ? mildstones[0] : mildstone ;
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
          item.OwnerTitle = item.AssignedTo && item.AssignedTo.length > 0  ? item.AssignedTo[0].Title : "";           

          // Time Line
          timeline.push({
            id: item.Project_x0020_ID,
            name: item.Project,
            start: item.StartDate,
            end: item.DueDate,
            progress: 10
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
          if(item.Team_x0020_Member){
          if(item.Team_x0020_Member.Email){
            item.Team_x0020_Member.ImgUrl =   "https://esplrms-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/" +
            item.Team_x0020_Member.Email.split("@")[0].toLowerCase() +
            "_esplrms_onmicrosoft_com_MThumb.jpg";
          }else{
            item.Team_x0020_Member.ImgUrl = "https://esplrms.sharepoint.com/sites/projects/SiteAssets/default.jpg";
          }
        }            
        });  
        let projects = this.state.projectList;
        let project = find(projects, {"Project" : name })
        project.TeamMemberList = response;  
        this.setState({ projectList: projects });    
      });
  }
  private getMildStonesByProject(name): void {
    let filter = "Project/Title eq '" + name + "' and Duration eq 0";
    sp.web.lists
      .getByTitle("Tasks List")
      .items.select("Title","StartDate", "DueDate","Body", "Status0/ID",  "Status0/Status", "Status0/Status_x0020_Color","Project/ID", "Project/Title", "AssignedTo/Title", "AssignedTo/ID",
      "AssignedTo/EMail","Priority")
      .expand("Project", "Status0","AssignedTo")
      .filter(filter)
      .get()
      .then((response: Array<MildStones>) => {
        response.forEach(item =>{
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
        let project = find(projects, {"Project" : name })
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
      let project = find(projects, {"Project" : name })
      project.DocumentList = response; 
      this.setState({ projectList: projects });
      });
  }

  private getTaggingByProject(name): void {
    let filter = "Project eq '" + name + "'";
    sp.web.lists
      .getByTitle("Project")
      .items
      .select("Project_x0020_Tag/ID","Project_x0020_Tag/Tags").expand("Project_x0020_Tag")
      .filter("Project eq 'AlphaServe'")
      .get()
      .then((response: Array<any>) => {
        if(response != null && response.length > 0){
          let tags = response[0].Project_x0020_Tag;
          let projects = this.state.projectList;
          let project = find(projects, {"Project" : name })
          project.TagList = tags; 
          this.setState({ projectList: projects });
        }
      });
  }
}
