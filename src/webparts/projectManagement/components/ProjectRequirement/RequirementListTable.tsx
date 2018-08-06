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
                <div style={{textAlign : "center"}}> 
                    {rowData.Attachments.toString()}
                </div>
            );
    }

    impactTemplate(rowData: Requirement, Impact_x0020_On_x0020_Timelines) {
        if (rowData.Impact_x0020_On_x0020_Timelines  === true){
            return (
                <div style={{textAlign : "center"}}>
                    Yes
                </div>
            );
        }
            else{
                return (
                    <div style={{textAlign : "center"}}>
                        No
                    </div>
                );
            }
    }

    ownerTemplate(rowData: Requirement, column) {
        if (rowData.Author)
            return (
                <div>
                    {rowData.Author.Title}
                </div>
            );
    }
    resourceTemplate(rowData: Requirement, column) {
        if (rowData.Resources)
            return (
                <div style={{textAlign : "center"}}>
                    {rowData.Resources}
                </div>
            );
    }
    effortsTemplate(rowData: Requirement, column) {
        if (rowData.Efforts)
            return (
                <div style={{textAlign : "center"}}>
                    {rowData.Efforts}
                </div>
            );
    }


    editTemplate(rowData, column) {
        return <a href="#" onClick={this.onEditProject.bind(this, rowData)}><i className="far fa-edit"></i></a>;
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    actionTemplate(rowData, column) {
        return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}><i className="fas fa-trash-alt" ></i></a>;
    }
    fileTemplate(rowData: Requirement, column) {
        if (rowData.AttachmentFiles) {
            let iconClass = "";
            let type = "";
             if(rowData.AttachmentFiles.length>0){
             let data = rowData.AttachmentFiles[0].FileName;
            return (
                <div>
                    <a href={rowData.AttachmentFiles[0].ServerRelativeUrl} >{rowData.AttachmentFiles[0].FileName} </a>  
                </div>
            );
        }
        }
    }

    RequirementTemplate(rowData: Requirement, column) {
        if (rowData.Requirement)
            return (
                // <div className={styles.Responsibility}>
                <div style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis"}}>
                    {/* {rowData.Roles_Responsibility} */}
                    <span title={rowData.Requirement}>{rowData.Requirement}</span>
                </div>
            );
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
        var result = confirm("Are you sure you want to delete item?");
        if (result) {
        e.preventDefault();
           console.log('Edit :' + rowData);
           
           
        
           sp.web.lists.getByTitle(this.props.list).
           items.getById(rowData.ID).delete().then((response) => {
             console.log(this.props.list + ` item deleted`);
             this.getScheduleList(this.props.list);
           });
      // this.getScheduleList(this.props.list);
         
       }
    }
    
    public render(): React.ReactElement<IRequirementState> {
        return (
            <div className="">
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <h5>Requirements</h5>
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px"}} onClick={this.onAddProject}>
                        Add Requirements
                    </button>
                    {this.state.showComponent ?
                        <AddRequirement id={this.state.projectID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} list={this.props.list} projectId={this.props.projectId} /> :
                        null
                    }
                    <div className="requirement-list">
                        <DataTable value={this.state.projectList} paginator={true} rows={5} responsive={true} rowsPerPageOptions={[5, 10, 20]}>
                            <Column body={this.editTemplate}style={{ width: "3%", textAlign:"center" }} />
                            <Column field="Requirement" sortable={true} header="Requirement"style={{ width: "29%" }} body={this.RequirementTemplate} />
                            <Column field="AttachmentFiles" header="Attachment"  body={this.fileTemplate}style={{ width: "29%" }} />
                            <Column field="Resources" sortable={true} header="Resources"style={{ width: "10%", align: "center" }} body={this.resourceTemplate}  />
                            
                             <Column field="Impact_x0020_On_x0020_Timelines" sortable={true} header="Impact on Timelines?"  body={this.impactTemplate} style={{ width: "20%" }} /> 
                            <Column field="Efforts" header="Efforts" sortable={true}style={{ width: "10%" }}body={this.effortsTemplate}   />
                           
                            {/* <Column field="Apporval_x0020_Status" sortable={true} header="Approval Status" />
                             <Column field="Approver" header="Approver" sortable={true} body={this.approverTemplate} />  */}
                            {/* <Column field="Author" header="Created By" sortable={true} body={this.ownerTemplate} />
                            <Column field="Created" header="Created On" sortable={true} body={this.duedateTemplate} /> */}
                            <Column  body={this.actionTemplate}style={{ width: "3%", textAlign:"center" }} />
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


            sp.web.lists.getByTitle(list).items.select("ID", "Requirement", "Resources", "Impact_x0020_On_x0020_Timelines", "Efforts", "Attachments", "Apporval_x0020_Status", "Approver/Title", "Approver/ID", "Author/Title", "Author/ID", "Created","AttachmentFiles","AttachmentFiles/ServerRelativeUrl","AttachmentFiles/FileName")
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



