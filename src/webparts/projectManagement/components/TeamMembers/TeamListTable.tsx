import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {ITeamMembersProps } from "./ITeamMembersProps";
import { ITeamState } from "./ITeamState";
import {
    TeamMembers
} from "./TeamList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';

export default class ProjectListTable extends React.Component<
ITeamMembersProps,
ITeamState
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
            projectList: new Array<TeamMembers>(),
            showComponent: false
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
    }
    refreshGrid (){
        this.getAllProjectMemeber()
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getAllProjectMemeber();
     
        
    }
    componentWillReceiveProps(nextProps) { }

    /* Private Methods */

    /* Html UI */
    duedateTemplate(rowData: TeamMembers, column) {
        if (rowData.Start_x0020_Date)
            return (
                <div>
                    {(new Date(rowData.Start_x0020_Date)).toLocaleDateString()}
                </div>
            );
    }
    enddateTemplate(rowData: TeamMembers, column) {
        if (rowData.End_x0020_Date)
            return (
                <div>
                    {(new Date(rowData.End_x0020_Date)).toLocaleDateString()}
                </div>
            );
    }

    

    ownerTemplate(rowData: TeamMembers, column) {
        if (rowData.Team_x0020_Member)
            return (
                <div>
                    {rowData.Team_x0020_Member.Title}
                </div>
            );
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
    public render(): React.ReactElement<ITeamState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Team Members
                    </button>
                    {this.state.showComponent ?
                            <AddProject parentMethod={this.refreshGrid}/>  :
                        null
                    }
                    <DataTable value={this.state.projectList} paginator={true} rows={10} rowsPerPageOptions={[5, 10, 20]}>
                        <Column field="AssignedTo" header="Owner" body={this.ownerTemplate} />
                       
                        <Column field="Start_x0020_Date" header="Start Date" body={this.duedateTemplate}  />
                        <Column field="End_x0020_Date" header="End Date" body={this.enddateTemplate} />
                        <Column field="Status" header="Status" />
                       
                        <Column header="Remove" body={this.actionTemplate} />
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    

    getAllProjectMemeber(){
        sp.web.lists.getByTitle("Project Team Members").items.select("Team_x0020_Member/ID", "Team_x0020_Member/Title","Start_x0020_Date", "End_x0020_Date","Status").expand("Team_x0020_Member").getAll().then((response) => {
            console.log('member by name', response);
            this.setState({ projectList: response });
            
        });
      }
     
    }