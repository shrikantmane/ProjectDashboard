import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {IScheduleProps } from "./IScheduleProps";
import { IScheduleState } from "./IScheduleState";
import {
    Schedule
} from "./ScheduleList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';

export default class ProjectListTable extends React.Component<
IScheduleProps,
IScheduleState
    > {
    constructor(props) {
        super(props);
        // this.state = {
        //    projectList: new Array<Project>(),
        //     //   projectTimeLine: new Array<ProjectTimeLine>(),
        //     projectName: null,
        //     ownerName: null,
        //     status: null,
        //     priority: null,
        //     isLoading: true,
        //     isTeamMemberLoaded: false,
        //     isKeyDocumentLoaded: false,
        //     isTagLoaded: false,
        //     expandedRowID: -1,
        //     expandedRows: []
        // };
        this.state = {
            projectList: new Array<Schedule>(),
            showComponent: false
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getproject();
     
        
    }
    refreshGrid (){
        this.getproject()
    }
    componentWillReceiveProps(nextProps) { }

    /* Private Methods */

    /* Html UI */
    duedateTemplate(rowData: Schedule, column) {
        if (rowData.StartDate)
            return (
                <div>
                    {(new Date(rowData.StartDate)).toLocaleDateString()}
                </div>
            );
    }
    enddateTemplate(rowData: Schedule, column) {
        if (rowData.DueDate)
            return (
                <div>
                    {(new Date(rowData.DueDate)).toLocaleDateString()}
                </div>
            );
    }

    

    ownerTemplate(rowData: Schedule, column) {
        if (rowData.AssignedTo)
            return (
                <div>
                    {rowData.AssignedTo[0].Title}
                </div>
            );
    }
    statusTemplate(rowData: Schedule, column) {
        if (rowData.Status0)
            return (<span style={{ color: rowData.Status0['Status_x0020_Color'] }}>{rowData.Status0['Status']}</span>);
    }
    actionTemplate(rowData, column) {
        return <a href="#"> Remove</a>;
    }
    editTemplate(rowData, column) {
        return <a href="#"> Edit </a>;
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    public render(): React.ReactElement<IScheduleState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Task
                    </button>
                    {this.state.showComponent ?
                        <AddProject parentMethod={this.refreshGrid}/> :
                        null
                    }
                    <DataTable value={this.state.projectList} paginator={true} rows={10} rowsPerPageOptions={[5, 10, 20]}>
                        <Column field="Title" header="Task Name"  />
                        <Column field="StartDate" header="Start Date" body={this.duedateTemplate}  />
                        <Column field="DueDate" header="Due Date" body={this.enddateTemplate} />
                        <Column field="AssignedTo" header="Owner" body={this.ownerTemplate} />
                        <Column field="Duration" header="Duration" />
                        <Column field="Status0" header="Status" body={this.statusTemplate} />
                        <Column field="Priority" header="Priority" />
                       
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    

    // getAllProjectMemeber(){
    //     sp.web.lists.getByTitle("Project Team Members").items.select("Team_x0020_Member/ID", "Team_x0020_Member/Title","Start_x0020_Date", "End_x0020_Date","Status").expand("Team_x0020_Member").getAll().then((response) => {
    //         console.log('member by name', response);
    //         this.setState({ projectList: response });
            
    //     });
    //   }
      getproject() {
          var scheduleList:string;
        // get Project Documents list items for all projects
        sp.web.lists.getByTitle("Project").items
          .select("Project", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "Priority","Task_x0020_List").expand("Status0", "AssignedTo")
          .filter("Project eq 'AlphaServe'")
          .getAll()
          .then((response) => {
            console.log('Project by names', response);
            scheduleList=response[0].Task_x0020_List;
            
            this.getScheduleList(scheduleList);
        }).catch((e: Error) => {
            alert(`There was an error : ${e.message}`);
          });
          
         
   
      }
     getScheduleList(ListName){
        sp.web.lists.getByTitle(ListName).items.select("Title", "StartDate","DueDate", "Duration","Priority","AssignedTo/Title", "AssignedTo/ID", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color")
        .expand("AssignedTo","Status0")
        .filter("Project/ID eq 1")
        .get().
        then((response) => {
            console.log('member by list', response);
            this.setState({ projectList: response });
            
        });
     }
      
     
    }