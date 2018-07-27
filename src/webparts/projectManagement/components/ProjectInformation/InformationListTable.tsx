import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {IInformationProps } from "./IInformationProps";

import { IInformationState } from "./IInformationState";
import {
     Information
} from "./InformationList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
//import AddProject from '../AddProject/AddProject';
import AddInformation from '../AddInformation/AddInformation';
export default class ProjectListTable extends React.Component<
IInformationProps,
IInformationState
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
            projectList: new Array<Information>(),
            showComponent: false
        };
        this.onAddProject = this.onAddProject.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getProjectInformation();
        this.refreshGrid = this.refreshGrid.bind(this);
     
        
    }
    refreshGrid (){
        this.getProjectInformation()
    }
    componentWillReceiveProps(nextProps) { }

    /* Private Methods */

    /* Html UI */
    

    
    // deptTemplate(rowData: Information, column) {
    //     if (rowData.Owner)
    //         return (
    //             <div>
    //                 {rowData.Owner.Department}
    //             </div>
    //         );
    // }

    ownerTemplate(rowData: Information, column) {
        if (rowData.Owner)
            return (
                <div>
                    {rowData.Owner.Title}
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
    public render(): React.ReactElement<IInformationState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Information
                    </button>
                    {this.state.showComponent ?
                         <AddInformation parentMethod={this.refreshGrid}/> :
                        null
                    }
                    <DataTable value={this.state.projectList} paginator={true} rows={10} rowsPerPageOptions={[5, 10, 20]}>
                    <Column header="Edit" body={this.editTemplate} />
                        <Column field="Roles_Responsibility" header="Role/Responsibility"  />
                       
                        <Column field="Owner" header="Owner" body={this.ownerTemplate}  />
                        <Column field="Department" header="Department"  /> 
                       
                        <Column header="Remove" body={this.actionTemplate} />
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

      getProjectInformation(){
        sp.web.lists.getByTitle("Project Information").items
          .select("Roles_Responsibility","Owner/ID","Owner/Title","Owner/FirstName","Owner/Department").expand("Owner/ID")
         .get()
       .then((response) => {
            console.log('member by name', response);
            this.setState({ projectList: response });
            
        });
      }
    //   private GetUserProperties(): void {  
    //     sp.profiles.myProperties.get().then(function(result) {  
    //         var userProperties = result.UserProfileProperties;  
    //         console.log("hello",userProperties);
    //         var userPropertyValues = "";  
    //         userProperties.forEach(function(property) {  
    //             userPropertyValues += property.Key + " - " + property.Value + "<br/>";  
    //         });  
    //         document.getElementById("spUserProfileProperties").innerHTML = userPropertyValues;  
    //     }).catch(function(error) {  
    //         console.log("Error: " + error);  
    //     });  
    // }  
    getprofile(){
    sp.profiles.getPropertiesFor("i:0#.f|membership|LeeG@esplrms.onmicrosoft.com").then(function(result) {
        var props = result.UserProfileProperties;
        var propValue = "";
        props.forEach(function(prop) {
            propValue += prop.Key + " - " + prop.Value + "<br/>";
        });
        
        console.log("hie",props[12])
    }).catch(function(err) {
        console.log("Error: " + err);
    });
}

 
    }