import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {IProjectRequirementProps } from "./IProjectRequirementProps";
import { IRequirementState } from "./IRequirementState";
import {
    Requirement
} from "./RequirementList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
//import AddProject from '../AddProject/AddProject';
import AddRequirement from '../AddRequirement/AddRequirement';
export default class ProjectListTable extends React.Component<
IProjectRequirementProps,
IRequirementState
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
            projectList: new Array<Requirement>(),
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
        this.GetRequirements();
     
        
    }
    refreshGrid (){
        this.GetRequirements();
    }
    componentWillReceiveProps(nextProps) { }

    /* Private Methods */

    /* Html UI */
    duedateTemplate(rowData: Requirement, column) {
        if (rowData.Created)
            return (
                <div>
                    {(new Date(rowData.Created)).toLocaleDateString()}
                </div>
            );
    }
    
   attachmentTemplate(rowData: Requirement, Attachments) {
        if (rowData.Attachments)
            return (
                <div>
                    {rowData.Attachments.toString()}
                </div>
            );
    }
    
    impactTemplate(rowData: Requirement, Impact_x0020_on_x0020_Timelines) {
        if (rowData.Impact_x0020_on_x0020_Timelines)
            return (
                <div>
                    {rowData.Impact_x0020_on_x0020_Timelines.toString()}
                </div>
            );
    }
    
    ownerTemplate(rowData: Requirement, column) {
        if (rowData.Author)
            return (
                <div>
                    {rowData.Author.Title}
                </div>
            );
    }
    approverTemplate(rowData: Requirement, column) {
        if (rowData.Approver)
            return (
                <div>
                    {rowData.Approver.Title}
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
    public render(): React.ReactElement<IRequirementState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Requirement
                    </button>
                    {this.state.showComponent ?
                         <AddRequirement parentMethod={this.refreshGrid}/> :
                        null
                    }
                    <DataTable value={this.state.projectList} paginator={true} rows={10} responsive={true} rowsPerPageOptions={[5, 10, 20]}>
                   
                        <Column field="Requirement" header="Requirement"  />
                        <Column field="Resources" header="Resources" />
                        <Column field="Impact_x0020_on_x0020_Timelines" header="Impact on Timeline?"   body={this.impactTemplate}  />
                        <Column field="Efforts" header="Efforts" sortable={true} />
                        <Column field="Attachments" header="Attachment"  body={this.attachmentTemplate} />
                        <Column field="Apporval_x0020_Status" header="Approval Status" />
                        <Column field="Approver" header="Approver"  body={this.approverTemplate}  />
                        <Column field="Author" header="Created By"  body={this.ownerTemplate}  />
                        <Column field="Created" header="Created On"  body={this.duedateTemplate}  />
                        
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    
   
    GetRequirements() {
        sp.web.lists.getByTitle("Requirements").items
          .select("Requirement", "Resources", "Impact_x0020_on_x0020_Timelines", "Efforts", "Attachments", "Apporval_x0020_Status", "Approver/Title", "Approver/ID", "Author/Title", "Author/ID", "Created").expand("Approver", "Author")
          .getAll().then((response) => {
            console.log('member by name', response);
            this.setState({ projectList: response });
          });
     
      }
     
     
     
}