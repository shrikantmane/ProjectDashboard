import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { IInformationProps } from "./IInformationProps";

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
            showComponent: false, 
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.reopenPanel = this.reopenPanel.bind(this);
        this.editTemplate = this.editTemplate.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
        this.actionTemplate = this.actionTemplate.bind(this);
        this.RoleTemplate = this.RoleTemplate.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if (this.props.list != "" || this.props.list != null) {
            this.getProjectInformation(this.props.list);
        }




    }

    refreshGrid() {
        this.setState({
            showComponent: false,
            informationID: null
        })
        this.getProjectInformation(this.props.list);
    }
    reopenPanel() {
        this.setState({
            showComponent: false,
            informationID: null
        })
    }
    componentWillReceiveProps(nextProps) {
        if (nextProps.list != "" || nextProps.list != null) {
            this.getProjectInformation(nextProps.list);
        }
    }

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

  
    actionTemplate(rowData, column) {
        return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}><i className="far fa-times-circle"></i></a>;
    }
    editTemplate(rowData, column) {
        return <a href="#" onClick={this.onEditProject.bind(this, rowData)}><i className="far fa-edit"></i> Edit</a>;
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    private onEditProject(rowData, e): any {
        e.preventDefault();
        console.log('Edit :' + rowData);
        this.setState({
            showComponent: true,
            informationID: rowData.ID
        });
    }

    private deleteListItem(rowData, e): any {
        e.preventDefault();
        console.log('Edit :' + rowData);
        sp.web.lists.getByTitle(this.props.list).
        items.getById(rowData.ID).delete().then((response) => {
          console.log(this.props.list + ` item deleted`);
          this.getProjectInformation(this.props.list);
        });
 
    }

  ownerTemplate(rowData: Information, column) {
        if (rowData.Owner)
            return (
                <div>
                    {rowData.Owner.Title}
                </div>
            );
    }
    
     RoleTemplate(rowData: Information, column) {
        if (rowData.Roles_Responsibility)
            return (
                // <div className={styles.Responsibility}>
                <div style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis"}}>
                    {rowData.Roles_Responsibility}
                </div>
            );
    }

    


    public render(): React.ReactElement<IInformationState> {
        console.log('this.state.projectList', this.state.projectList)
        return (
            <div className="PanelContainer">
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <h5>Responsibilities</h5>
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px", marginLeft : "35%" }} onClick={this.onAddProject}>
                        Add Role/Responsibility
                    </button>
                    {this.state.showComponent ?
                        <AddInformation id={this.state.informationID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} list={this.props.list} projectId={this.props.projectId} /> :
                        null
                    }
                    <div className="responsibility-list">
                        <DataTable value={this.state.projectList} responsive={true} paginator={true} rows={5} rowsPerPageOptions={[5, 10, 20]}>
                            <Column header="Edit" body={this.editTemplate} style={{width: "15%"}}/>
                            <Column field="Roles_Responsibility" sortable={true} header="Responsibility"  body={this.RoleTemplate} />
                            <Column field="OwnerName" sortable={true} header="Owner" style={{width: "22%"}}   />
                            <Column field="Department" sortable={true} header="Department" style={{width: "25%"}}/>
                            <Column header="" body={this.actionTemplate} style={{width: "7%"}}/>
                        </DataTable>
                    </div>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    // getProjectInformation(list) {
      
    //     if ((list) != "") {
    //         sp.web.lists.getByTitle(list).items
    //             .select("ID", "Roles_Responsibility", "Owner/ID", "Owner/Title","Owner/EMail").expand("Owner")
    //             .get()
    //             .then((response) => {
    //                 console.log('infor by name', response);
    //               // this.setState({ projectList: response });
    //                 for (var i = 0; i < response.length; i++) {
    //                     var Id = response[i].ID;
    //                  //   var Title = response[i].Owner.Title;
    //                   //  var Email = response[i].Owner.EMail;
    //                    var Responsibility=response[i].Roles_Responsibility;
    //                    var owner=response[i].Owner;
    //                     this.GetUserProperties(owner,Responsibility,Id);
    //                   }

    //             });
    //     }
    // }


    private getProjectInformation(list) {
        if ((list) != "") {
        sp.web.lists.getByTitle(list).items
        .select("ID", "Roles_Responsibility", "Owner/ID", "Owner/Title","Owner/EMail").expand("Owner")
          .get()
          .then((response: Array<Information>) => {
            let currentScope = this;
            let count = 1;
            response.forEach(item => {
              if (item.Owner) {
                let loginName = "i:0#.f|membership|" + item.Owner.EMail;
                sp.profiles.getPropertiesFor(loginName).then(function (result) {
                  item.OwnerName = item.Owner.Title;
                  item.Department = result.UserProfileProperties[13].Value;
                  item.Owner.Department = result.UserProfileProperties[13].Value;                
                  if (count == response.length) {
                    currentScope.setState({ projectList: response });
                  }
                  count++;
                });
              } else {
                if (count == response.length) {
                  currentScope.setState({ projectList: response });
                }
                count++;
              }
            });
          });
        }
      }
    // private GetUserProperties(owner,Responsibility,Id) {
      
    //     let loginName = "i:0#.f|membership|" + owner.EMail;
    //     let department1;
    //     let jobTitle;
    //     let pic;
    //     sp.profiles.getPropertiesFor(loginName).then((response) =>  {
    //       console.log("result.UserProfileProperties -", response.UserProfileProperties);
         
    //       department1 = response.UserProfileProperties[13].Value;
    //       pic = response.UserProfileProperties[16].Value;
    //       jobTitle = response.UserProfileProperties[21].Value;
    //       console.log("department11",department1)
          
    //       console.log('prevState.projectList', this.state.projectList);
    //      this.setState(prevState => ({
    //         projectList: [...prevState.projectList, {ID:Id, Roles_Responsibility: Responsibility, Owner: owner.Title,Department:department1  }]
    //        }));
    
    //     }).catch(function (err) {
    //       console.log("Error: " + err);
    //     });
    
    //   }
   


}