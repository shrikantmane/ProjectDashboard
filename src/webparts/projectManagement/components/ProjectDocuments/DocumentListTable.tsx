import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {IProjectDocumentProps } from "./IProjectDocumentProps";
import { IDocumentState } from "./IDocumentState";
import {
    Document
} from "./DocumentList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';

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
            selectedFile:"",
            documentID:""
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
        this.UploadFiles=this.UploadFiles.bind(this);
        this.deleteListItem=this.deleteListItem.bind(this);
       
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if(this.props.list!="" || this.props.list!=null){
            this.getProjectDocuments(this.props.list);
        }
     
        
    }
    refreshGrid (){
        this.getProjectDocuments(this.props.list);
    }
   
     deleteListItem(listName, itemID):any {
        sp.web.lists.getByTitle(listName).
          items.getById(itemID).delete().then(i => {
            console.log(listName + ` item deleted`);
          });
      }
    componentWillReceiveProps(nextProps) {
        if(nextProps.list!="" || nextProps.list!=null){
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
        if (rowData.File)
       {
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
                <div>

                 <a href={rowData.File.ServerRelativeUrl} >{rowData.File.Name} </a> 
                 <i
                 style={{ marginRight: "5px" }}
                                className={iconClass}/>
                </div>
            );
    
}}


    
    actionTemplate(rowData, column) {
        return <a href="#"   > Remove</a>;
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

    UploadFiles() {
       
       //in case of multiple files,iterate or else upload the first file.
       
        var file =this.state.selectedFile;
        if (file!=undefined || file!=null){
    
        //assuming that the name of document library is Documents, change as per your requirement, 
        //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
        sp.web.getFolderByServerRelativeUrl("/projects/Project Documents").files.add(file.name, file, true).then((result) => {
            console.log(file.name + " upload successfully!");
        })
    }
    
        }
        fileChangedHandler = (event) => {
            const file = event.target.files[0]
            this.setState({selectedFile: event.target.files[0]})
          }
    public render(): React.ReactElement<IDocumentState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
              
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Document
                    </button>
                   
                     {this.state.showComponent ?
                        <input type="file"  onChange={this.fileChangedHandler} />:
                        null
                    } 
                     {this.state.showComponent ?
                       <input type="button" value="Upload"  onClick={this.UploadFiles} />:
                        null
                    } 
                    <DataTable value={this.state.projectList} paginator={true} rows={10} responsive={true} rowsPerPageOptions={[5, 10, 20]}>
                   
                        <Column field="Title" header="Attachment" body={this.fileTemplate} 
                        
                        />
                        
                        <Column field="Author" header="Created By"  body={this.ownerTemplate}  />
                        <Column field="Created" header="Created On"  body={this.duedateTemplate}  />
                        <Column header="Remove" body={this.actionTemplate} />
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    
   
   

      getProjectDocuments(list){
        if((list)!=""){
        
        sp.web.lists.getByTitle(list).items
          .select("ID","File","Author/ID","Author/Title","Created").expand("File","Author")
         
         .get()
         .then((response) => {
            console.log('member by name', response);
            this.setState({ projectList: response });
          });
        }
      }
     
     
     
     
}