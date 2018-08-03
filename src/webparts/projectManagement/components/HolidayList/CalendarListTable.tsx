import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { ICalendarProps } from "./ICalendarProps";

import { ICalendarState } from "./ICalendarState";
import {
    CalendarList
} from "./CalendarList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
//import AddProject from '../AddProject/AddProject';
import AddEvent from '../AddEvent/AddEvent';
export default class ProjectListTable extends React.Component<
ICalendarProps,
    ICalendarState
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
            projectList: new Array<CalendarList>(),
            showComponent: false,
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.reopenPanel = this.reopenPanel.bind(this);
        this.editTemplate = this.editTemplate.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
        this.actionTemplate = this.actionTemplate.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if (this.props.list != "" || this.props.list != null) {
            this.getProjectCalendar(this.props.list);
        }




    }

    refreshGrid() {
        this.setState({
            showComponent: false,
            informationID: null
        })
        this.props.onRefreshCalender();
        this.getProjectCalendar(this.props.list);
    }
    reopenPanel() {
        this.setState({
            showComponent: false,
            informationID: null
        })
    }
    componentWillReceiveProps(nextProps) {
        if (nextProps.list != "" || nextProps.list != null) {
            this.getProjectCalendar(nextProps.list);
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

    // ownerTemplate(rowData: Information, column) {
    //     if (rowData.Owner)
    //         return (
    //             <div>
    //                 {rowData.Owner.Title}
    //             </div>
    //         );
    // }
    actionTemplate(rowData, column) {
        return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}><i className="fas fa-trash-alt"></i></a>;
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
    duedateTemplate(rowData:CalendarList , column) {
        if (rowData.EndDate)
            return (
                <div>
                    {(new Date(rowData.EndDate)).toLocaleDateString()}
                </div>
            );
    }
    startdateTemplate(rowData:CalendarList , column) {
        if (rowData.EventDate)
            return (
                <div>
                    {(new Date(rowData.EventDate)).toLocaleDateString()}
                </div>
            );
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
          this.getProjectCalendar(this.props.list);
        });
 
    }
    public render(): React.ReactElement<ICalendarState> {
        return (
            <div className="PanelContainer">
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <h5>Holidays</h5>
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Event
                    </button>
                    {this.state.showComponent ?
                        <AddEvent id={this.state.informationID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} list={this.props.list} projectId={this.props.projectId}  /> :
                        null
                    }
                    <div className="holiday-list">
                        <DataTable value={this.state.projectList} responsive={true} paginator={true} rows={5} rowsPerPageOptions={[5, 10, 20]}>
                            <Column header="Edit" body={this.editTemplate} style={{width: "15%"}}/>
                            <Column header="Title" field="Title" />
                            <Column field="EventDate" sortable={true} header="Start Date"   body={this.startdateTemplate}  style={{width: "21%"}}/>
                            {/* <Column field="Owner" sortable={true} header="Owner" body={this.ownerTemplate} /> */}
                            <Column field="EndDate" sortable={true} header="End Date" body={this.duedateTemplate}  style={{width: "20%"}}/>
                            <Column header="" body={this.actionTemplate} style={{width: "7%"}} />
                        </DataTable>
                    </div>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    getProjectCalendar(list) {
        if ((list) != "") {
            sp.web.lists.getByTitle(list).items
                .select("ID", "Title","EventDate","EndDate")
                .get()
                .then((response) => {
                    console.log('infor by name', response);
                    this.setState({ projectList: response });

                });
        }
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



}