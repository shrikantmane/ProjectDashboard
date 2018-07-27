import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {IProjectRiskProps } from "./IProjectRiskProps";
import { IRiskState } from "./IRiskState";
import {
    Risk
} from "./RiskList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
//import AddProject from '../AddProject/AddProject';
import AddRisk from '../AddRisk/AddRisk';
export default class ProjectListTable extends React.Component<
IProjectRiskProps,
IRiskState
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
            projectList: new Array<Risk>(),
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
        this.GetRisk();
     
        
    }
    refreshGrid (){
        this.GetRisk();
    }
    componentWillReceiveProps(nextProps) { }

    /* Private Methods */

    /* Html UI */
    duedateTemplate(rowData: Risk, column) {
        if (rowData.Created)
            return (
                <div>
                    {(new Date(rowData.Created)).toLocaleDateString()}
                </div>
            );
    }
    
    targetdateTemplate(rowData: Risk, column) {
        if (rowData.Target_x0020_Date)
            return (
                <div>
                    {(new Date(rowData.Target_x0020_Date)).toLocaleDateString()}
                </div>
            );
    }
    

    ownerTemplate(rowData: Risk, column) {
        if (rowData.Author)
            return (
                <div>
                    {rowData.Author.Title}
                </div>
            );
    }
    assignedtoTemplate(rowData: Risk, column) {
        if (rowData.Assigned_x0020_To)
            return (
                <div>
                    {rowData.Assigned_x0020_To.Title}
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
    public render(): React.ReactElement<IRiskState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Risk
                    </button>
                    {this.state.showComponent ?
                         <AddRisk parentMethod={this.refreshGrid}/> :
                        null
                    }
                    <DataTable value={this.state.projectList} paginator={true} rows={10} responsive={true} rowsPerPageOptions={[5, 10, 20]}>
                    <Column header="Action" body={this.editTemplate} />
                        <Column field="Risk" header="Risk"  />
                        <Column field="AssignedTo" header="Owner" body={this.ownerTemplate} />
                        <Column field="Impact" header="Impact" />
                        <Column field="Target_x0020_Date" header="Target Date" body={this.targetdateTemplate} sortable={true} />
                        <Column field="Mitigation" header="Mitigation" />
                        <Column field="Author" header="Created By" body={this.ownerTemplate} />
                        <Column field="Created" header="Created On" sortable={true} body={this.duedateTemplate}  />
                       
                        <Column header="Remove" body={this.actionTemplate} />
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    
    GetRisk() {
        // get Risks list items for all projects
        sp.web.lists.getByTitle("Risks").items
          .select("Risk", "Impact", "Target_x0020_Date", "Mitigation","Author/Title", "Author/ID", "Created", "Assigned_x0020_To/Title", "Assigned_x0020_To/ID").expand("Author  ", "Assigned_x0020_To")
          .getAll().then((response) => {
            console.log('member by name', response);
            this.setState({ projectList: response });
     
      });
    }
     
     
}