import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { IProjectListProps } from "./IProjectListProps";
import { IProjectListState } from "./IProjectState";
import {
    Project
} from "./ProjectList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";

export default class ProjectListTable extends React.Component<
    IProjectListProps,
    IProjectListState
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
            projectList: new Array<Project>()
        };
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getProjectList();
    }
    componentWillReceiveProps(nextProps) { }

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
            return (<span style={{ color: rowData.Status0['Status_x0020_Color'] }}>{rowData.Status0['Status']}</span>);
    }

    ownerTemplate(rowData: Project, column) {
        if (rowData.AssignedTo)
            return (
                <div>
                    {rowData.AssignedTo[0].Title}
                </div>
            );
    }
    actionTemplate(rowData, column) {
        return <a href="#"> View Details </a>;
    }
    editTemplate(rowData, column) {
        return <a href="#"> Edit </a>;
    }
    public render(): React.ReactElement<IProjectListProps> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}

                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }}>
                        Add Project
                    </button>
                    <DataTable value={this.state.projectList} paginator={true} rows={10} rowsPerPageOptions={[5, 10, 20]}>
                        <Column header="Action" body={this.editTemplate} />
                        <Column field="Project" header="Project" />
                        <Column field="DueDate" header="Due Date" body={this.duedateTemplate} />
                        <Column field="Status0" header="Status" body={this.statusTemplate} />
                        <Column field="AssignedTo" header="Owner" body={this.ownerTemplate} />
                        <Column field="Priority" header="Priority" />
                        <Column field="Tag" header="Tags" />
                        <Column header="Project Details" body={this.actionTemplate} />
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
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
                this.getprojecttag();
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });

    }

    getprojecttag() {
        sp.web.lists.getByTitle("Project Tags").items
            .select("Project/Title", "Project/ID", "Tag").expand("Project")
            .get()
            .then((response) => {
                console.log(' all Project tag -', response);
                if (response != null && response.length > 0) {
                    let projects: any = this.state.projectList;
                    projects.forEach(element => {
                        let tagData: any = find(response, function (o) { return o.Project.Title === element.Project; })
                        let tags = tagData ? tagData.Tag : '';
                        console.log(tags);
                        element.Tag = tags;
                    });
                    this.setState({ projectList: projects });
                }
            });
    }



}
