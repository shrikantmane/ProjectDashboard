import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { IProjectRequirementProps } from "./IProjectRequirementProps";
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
        this.actionTemplate=this.actionTemplate.bind(this);
        this.reopenPanel = this.reopenPanel.bind(this);
        this.editTemplate = this.editTemplate.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if (this.props.list != "" || this.props.list != null) {
            this.getScheduleList(this.props.list);
        }


    }
    reopenPanel() {
        this.setState({
            showComponent: false,
            projectID: null
        })
    }

    refreshGrid() {
        this.setState({
            showComponent: false,
            projectID: null
        })
        this.getScheduleList(this.props.list);

    }
    componentWillReceiveProps(nextProps) {
        if (nextProps.list != "" || nextProps.list != null) {
            this.getScheduleList(nextProps.list);
        }
    }

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



    editTemplate(rowData, column) {
        return <a href="#" onClick={this.onEditProject.bind(this, rowData)}> Edit </a>;
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    actionTemplate(rowData, column) {
        return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}> Remove</a>;
    }
    fileTemplate(rowData: Requirement, column) {
        if (rowData.AttachmentFiles) {
            let iconClass = "";
            let type = "";
            if(rowData.AttachmentFiles[0].length>0){
            let data = rowData.AttachmentFiles[0].FileName;}
            // if (data.length > 1) {
            //     type = data[1];
            // }
            // switch (type.toLowerCase()) {
            //     case "doc":
            //     case "docx":
            //         iconClass = "far fa-file-word";
            //         break;
            //     case "pdf":
            //         iconClass = "far fa-file-pdf";
            //         break;
            //     case "xls":
            //     case "xlsx":
            //         iconClass = "far fa-file-excel";
            //         break;
            //     case "png":
            //     case "jpeg":
            //     case "gif":
            //         iconClass = "far fa-file-image";
            //         break;
            //     default:
            //         iconClass = "fa fa-file";
            //         break;
            // }


            return (
                <div>

                    <a href={rowData.AttachmentFiles[0].ServerRelativeUrl} >{rowData.AttachmentFiles[0].FileName} </a>
                    
                </div>
            );

        }
    }
    private onEditProject(rowData, e): any {
        e.preventDefault();
        console.log('Edit :' + rowData);
        this.setState({
            showComponent: true,
            projectID: rowData.ID
        });
    }
    private deleteListItem(rowData,e):any {
        e.preventDefault();
           console.log('Edit :' + rowData);
           
           
        
           sp.web.lists.getByTitle(this.props.list).
           items.getById(rowData.ID).delete().then((response) => {
             console.log(this.props.list + ` item deleted`);
           });
       this.getScheduleList(this.props.list);
         
       }
    
    public render(): React.ReactElement<IRequirementState> {
        return (
            <div className="PanelContainer">
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <h5>Requirements</h5>
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px"}} onClick={this.onAddProject}>
                        Add Requirement
                    </button>
                    {this.state.showComponent ?
                        <AddRequirement id={this.state.projectID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} list={this.props.list} projectId={this.props.projectId} /> :
                        null
                    }
                    <div className="requirement-list">
                        <DataTable value={this.state.projectList} paginator={true} rows={5} responsive={true} rowsPerPageOptions={[5, 10, 20]}>
                            <Column header="Edit" body={this.editTemplate} />
                            <Column field="Requirement" sortable={true} header="Requirement" />
                            <Column field="Resources" sortable={true} header="Resources" />
                             <Column field="Impact_x0020_on_x0020_Timelines" sortable={true} header="Impact on Timeliness?" body={this.impactTemplate} /> 
                            <Column field="Efforts" header="Efforts" sortable={true} />
                            <Column field="AttachmentFiles" header="Attachment" sortable={true} body={this.attachmentTemplate} />
                            {/* <Column field="Apporval_x0020_Status" sortable={true} header="Approval Status" />
                             <Column field="Approver" header="Approver" sortable={true} body={this.approverTemplate} />  */}
                            <Column field="Author" header="Created By" sortable={true} body={this.ownerTemplate} />
                            <Column field="Created" header="Created On" sortable={true} body={this.duedateTemplate} />
                            <Column header="Remove" body={this.actionTemplate} />
                        </DataTable>
                    </div>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/




    ////////////
    //   getproject() {
    //     var teamList:string;

    //     var last_part:string;
    //      var url =window.location.href ;
    //       var parts = url.split("/");
    //      last_part = parts[parts.length-1];

    //         var projectid=Number(last_part);
    //   // get Project Documents list items for all projects
    //   sp.web.lists.getByTitle("Project").items
    //     .select("Project", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "Priority","Task_x0020_List","Project_x0020_Team_x0020_Members","Project_x0020_Document","Requirements","ID").expand("Status0", "AssignedTo")
    //     .filter('ID eq \'' + projectid + '\'')
    //     .getAll()
    //     .then((response) => {
    //       console.log('Project by names', response);
    //       teamList=response[0].Requirements;

    //       this.getScheduleList(teamList,projectid);
    //   }).catch((e: Error) => {
    //       alert(`There was an error : ${e.message}`);
    //     });



    // }
    getScheduleList(list) {
        if ((list) != "") {


            sp.web.lists.getByTitle(list).items.select("ID", "Requirement", "Resources", "Impact_x0020_on_x0020_Timelines", "Efforts", "Attachments", "Apporval_x0020_Status", "Approver/Title", "Approver/ID", "Author/Title", "Author/ID", "Created","AttachmentFiles","AttachmentFiles/ServerRelativeUrl","AttachmentFiles/FileName")
                .expand("Approver", "Author","AttachmentFiles")
                .get().
                then((response) => {
                    console.log('requiremnts by list', response);
                   
                    this.setState({ projectList: response });
                 
                    
               // }
                });
            }
        }
    }



