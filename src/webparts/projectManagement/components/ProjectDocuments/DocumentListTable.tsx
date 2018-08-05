import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { IProjectDocumentProps } from "./IProjectDocumentProps";
import { IDocumentState } from "./IDocumentState";
import {
    Document
} from "./DocumentList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';
import AddDocument from '../AddDocument/AddDocument';
export default class ProjectListTable extends React.Component<
    IProjectDocumentProps,
    IDocumentState
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
            projectList: new Array<Document>(),
            showComponent: false,
            selectedFile: "",
            documentID: ""
        };
       
        this.refreshGrid = this.refreshGrid.bind(this);
        this.UploadFiles=this.UploadFiles.bind(this);
        this.actionTemplate=this.actionTemplate.bind(this);
        this.reopenPanel=this.reopenPanel.bind(this);
        this.onAddProject=this.onAddProject.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if (this.props.list != "" || this.props.list != null) {
            this.getProjectDocuments(this.props.list);
        }


    }
    refreshGrid() {
        this.getProjectDocuments(this.props.list);
    }

   
    private deleteListItem(rowData,e):any {
        var result = confirm("Are you sure you want to delete item?");
        if (result) {
        e.preventDefault();
           console.log('Edit :' + rowData);
           sp.web.lists.getByTitle(this.props.list).
           items.getById(rowData.ID).delete().then((response) => {
             console.log(this.props.list + ` item deleted`);
             this.getProjectDocuments(this.props.list);
           });
       }
    }
    componentWillReceiveProps(nextProps) {
        if (nextProps.list != "" || nextProps.list != null) {
            this.getProjectDocuments(nextProps.list);
        }
    }

    /* Private Methods */

    /* Html UI */
    duedateTemplate(rowData: Document, column) {
        if (rowData.Created)
            return (
                <div>
                    {(new Date(rowData.Created)).toLocaleDateString()}
                </div>
            );
    }




    ownerTemplate(rowData: Document, column) {
        if (rowData.Author)
            return (
                <div>
                    {rowData.Author.Title}
                </div>
            );
    }

    fileTemplate(rowData: Document, column) {
        if (rowData.File) {
            let iconClass = "";
            let type = "";
            let data = rowData.File.Name.split(".");
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
                <div  style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis"}}>

                    <a href={rowData.File.ServerRelativeUrl} ><i
                        style={{ marginRight: "5px" }}
                        className={iconClass} ></i> {rowData.File.Name} </a>
                    
                </div>
            );

        }
    }



    actionTemplate(rowData, column) {
        return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}><i className="fas fa-trash-alt"></i></a>;
    }
    editTemplate(rowData, column) {
        return <a href="#"> Edit </a>;
    }
   

    UploadFiles() {

        //in case of multiple files,iterate or else upload the first file.

        var file = this.state.selectedFile;
        if (file != undefined || file != null) {
            if (!this.state.selectedFile || this.state.selectedFile.length === 0) {
            
                this.setState({
                    showComponent: true,
                })
            }
            else{
                this.setState({
                    showComponent: false,
                })
            }
            //assuming that the name of document library is Documents, change as per your requirement, 
            //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
            sp.web.getFolderByServerRelativeUrl(this.props.list).files.add(file.name, file, true).then((result) => {
                console.log(file.name + " upload successfully!");
                this.setState({
                    selectedFile :""
                  });
                this.getProjectDocuments(this.props.list);
    

            });

        }
    }
       reopenPanel() {
        this.setState({
            showComponent: false,
            
        })
    }
    fileChangedHandler = (event) => {
        const file = event.target.files[0]
        this.setState({ selectedFile: event.target.files[0] })
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    public render(): React.ReactElement<IDocumentState> {
        return (
            <div className="">
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <h5>Documents</h5>
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Documents
                    </button>
                    {this.state.showComponent ?
                        <AddDocument parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} list={this.props.list} projectId={this.props.projectId} /> :
                        null
                    }
                    {/* <input type="button" value="Upload" className="btn btn-outline btn-sm" onClick={this.UploadFiles} /> */}
                    <div className="document-list">
                        <DataTable value={this.state.projectList} paginator={true} rows={5} responsive={true} rowsPerPageOptions={[5, 10, 20]}>
                            <Column field="Title" sortable={true} header="Documents" body={this.fileTemplate}style={{ width: "47%" }}  />
                            <Column field="Author" sortable={true} header="Created By" body={this.ownerTemplate}style={{ width: "23%" }} />
                            <Column field="Created" sortable={true} header="Created On" body={this.duedateTemplate}style={{ width: "23%" }} />
                            <Column  body={this.actionTemplate}style={{ width: "7%" }}  />
                        </DataTable>
                    </div>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/





    getProjectDocuments(list) {
        if ((list) != "") {

            sp.web.lists.getByTitle(list).items
                .select("ID", "File", "Author/ID", "Author/Title", "Created").expand("File", "Author")

                .get()
                .then((response) => {
                    console.log('member by name', response);
                    this.setState({ projectList: response });
                });
        }
    }




}