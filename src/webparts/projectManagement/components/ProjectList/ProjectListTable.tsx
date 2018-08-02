import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";

import { Link, Redirect } from 'react-router-dom';
import styles from "../ProjectManagement.module.scss";
import { IProjectListProps } from "./IProjectListProps";
import { IProjectListState } from "./IProjectState";
import {
    Project
} from "./ProjectList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';
import "bootstrap/dist/css/bootstrap.min.css";


export default class ProjectListTable extends React.Component<
    IProjectListProps,
    IProjectListState
    > {
    constructor(props) {
        super(props);
        this.state = {
            projectList: new Array<Project>(),
            showComponent: false
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
        this.reopenPanel = this.reopenPanel.bind(this);
        this.editTemplate = this.editTemplate.bind(this);
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getProjectList();
    }
    componentWillReceiveProps(nextProps) {

    }

    /* Private Methods */

    /* Html UI */
    duedateTemplate(rowData: Project, column) {
        if (rowData.DueDate)
            return (
                <div>
                    {(new Date(rowData.DueDate)).toLocaleDateString()}
                </div>
            );
    }

    statusTemplate(rowData: Project, column) {
        if (rowData.Status0)
            return (<div style={{ backgroundColor: rowData.Status0['Status_x0020_Color'], borderRadius: '3rem', textAlign: 'center', padding: '2px', color: 'white' }}>{rowData.Status0['Status']}</div>);
    }

    ownerTemplate(rowData: Project, column) {
        if (rowData.AssignedTo)
            return (
                <div>
                    {
                        rowData.AssignedTo.map((obj) =>
                            <span>{obj.Title}</span>
                        )
                    }
                </div>
            );
    }
    actionTemplate(rowData, column) {
        return (
            <div>
                <Link to={`/viewProjectDetails/${rowData.ID}`}><button className="btn btn-primary action-btn-style blue-light-color btn-xs" type="button"><abbr className="tooltip-style" title="Add member"><i className="fas fa-user-plus"></i></abbr></button></Link>
                <Link to={`/viewProjectDetails/${rowData.ID}`}><button className="btn  btn-primary action-btn-style blue-color btn-xs" type="button"><abbr className="tooltip-style" title="Add Document"><i className="fas fa-file"></i></abbr></button></Link>
                <Link to={`/viewProjectDetails/${rowData.ID}`}><button className="btn  btn-primary action-btn-style blue-dark-color btn-xs" type="button"><abbr className="tooltip-style" title="Requirments"><i className="fas fa-clipboard-list"></i></abbr></button></Link>
                <Link to={`/viewProjectDetails/${rowData.ID}`}><button className="btn btn-primary action-btn-style  black-color btn-xs" type="button"><abbr className="tooltip-style" title="View Details"><i className="white fas fa-info"></i></abbr></button></Link>
            </div>
        );
    }
    editTemplate(rowData, column) {
        return <a href="#" onClick={this.onEditProject.bind(this, rowData)}><i className="far fa-edit"></i> Edit </a>;
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
            projectID: rowData.ID
        });
    }
    refreshGrid() {
        this.setState({
            showComponent: false,
            projectID: null
        })
        this.getProjectList()
    }
    reopenPanel() {
        this.setState({
            showComponent: false,
            projectID: null
        })
    }
    public render(): React.ReactElement<IProjectListProps> {

        return (
            <div>
                {/* <DataTableSubmenu /> */}
                <div className="PanelContainer">
                    <div className="content-section implementation">
                        <h5>Projects</h5>
                        <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                            Add Project
                        </button>
                        {this.state.showComponent ?
                            <AddProject id={this.state.projectID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} /> :
                            null
                        }
                        <div className="project-list">
                            <DataTable value={this.state.projectList} responsive={true} paginator={true} rows={10} rowsPerPageOptions={[5, 10, 20]}>
                                <Column header="Action" body={this.editTemplate} style={{ width: "8%" }} />
                                <Column field="Project" sortable={true} header="Project" />
                                <Column field="DueDate" sortable={true} header="Due Date" body={this.duedateTemplate} />
                                <Column field="Status0" sortable={true} header="Status" body={this.statusTemplate} />
                                <Column field="AssignedTo" sortable={true} header="Owner" body={this.ownerTemplate} />
                                <Column field="Priority" sortable={true} header="Priority" />
                                <Column field="Tag" header="Tags" />
                                <Column header="Project Details" body={this.actionTemplate} />
                            </DataTable>
                        </div>
                    </div>
                </div>
                {/* <DataTableDoc></DataTableDoc> */}
            </div >
        );
    }

    /* Api Call*/

    private getProjectList(): void {
        // get Project Documents list items for all projects
        sp.web.lists.getByTitle("Project").items
            .select(
            "Project",
            "DueDate",
            "Status0/ID",
            "Status0/Status",
            "Status0/Status_x0020_Color",
            "AssignedTo/Title",
            "AssignedTo/ID",
            "Priority",
            "ID")
            .expand("Status0", "AssignedTo")
            .getAll()
            .then((response) => {
                console.log('Project by name', response);
                this.setState({ projectList: response });
                this.getProjectTag();
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });

    }

    getProjectTag() {
        sp.web.lists.getByTitle("Project Tags").items
            .select("Projects/ID", "Tag").expand("Projects")
            .get()
            .then((response) => {
                console.log(' all Project tag -', response);
                if (response != null && response.length > 0) {
                    let projects: any = this.state.projectList;
                    projects.forEach(item => {
                        let tagString = '';
                        response.forEach(element => {
                            if (element.Projects !== undefined && element.Projects.length > 0) {
                                let tempList = find(element.Projects, function (o: any) { return o.ID === item.ID; })
                                if (tempList && (element.Tag !== null || element.Tag !== '')) {
                                    tagString += element.Tag + ', ';
                                }
                            }
                        });
                        item.Tag = tagString.replace(/,\s*$/, "");
                    });
                    // response.forEach(element => {
                    //     if (element.Projects !== undefined && element.Projects.length > 0) {
                    //         //let tempList = find(element.Projects, function (o) { return o.ID === ; })

                    //         element.Projects.forEach(item => {

                    //             let tagData: any = find(projects, function (o: any) { return o.ID === item.ID; })
                    //             let tags = tagData ? tagData.Tag : '';
                    //             console.log(tags);
                    //             item.Tag = tags;
                    //         });
                    //     }
                    //     // let tagData: any = find(response, function (o) { return o.Project.ID === element.ID; })
                    //     // let tags = tagData ? tagData.Tag : '';
                    //     // console.log(tags);
                    //     // element.Tag = tags;
                    // });
                    this.setState({ projectList: projects });
                }
            });
    }



}
