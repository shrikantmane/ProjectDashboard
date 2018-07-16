import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";

import { ICEOProjectProps } from "./ICEOProjectProps";
import { ICEOProjectState } from "./ICEOProjectState";
import { CEOProjects } from "./CEOProject";
import CEOProjectTimeLine from "../CEOProjectTimeLine/CEOProjectTimeLine";
import ProjectTimeLine from "../CEOProjectTimeLine/ProjectTimeLine";
import { filter } from "lodash";

export default class CEOProjectTable extends React.Component<
  ICEOProjectProps,
  ICEOProjectState
> {
  constructor(props) {
    super(props);
    this.state = {
      projectList: new Array<CEOProjects>(),
      projectTimeLine: new Array<ProjectTimeLine>()
      //expandedRows:[]
    };
  }

  componentDidMount() {
    this.getProjectList();
  }
  componentWillReceiveProps(nextProps) {}

  /* Private Methods */

  private statusTemplate(rowData: CEOProjects, column) {
    if (rowData.Status0)
      return (
        <div
          style={{
            backgroundColor: rowData.Status0.Status_x0020_Color,
            height: "2.9em",
            width: "100%",
            textAlign: "center",
            paddingTop: 7,
            color: "#fff"
          }}
        >
          {rowData.Status0.Status}
        </div>
      );
  }

  private ownerTemplate(rowData: CEOProjects, column) {
    if (rowData.AssignedTo)
      return (
        <div>
          <img
            src={rowData.AssignedTo[0].imgURL}
            style={{ marginRight: "5px", width: "20px" }}
          />
          {rowData.AssignedTo[0].Title}
        </div>
      );
  }

  private rowExpansionTemplate(data) {
    return <div>test</div>;
  }
  /* Html UI */

  public render(): React.ReactElement<ICEOProjectProps> {
    return (
      <div>
        {this.state.projectTimeLine.length > 0 ? (
          <CEOProjectTimeLine tasks={this.state.projectTimeLine} />
        ) : null}
        <div style={{ marginTop: "10px" }}>
          <DataTable
            value={this.state.projectList}
            responsive={true}
            expandedRows={this.state.expandedRows}
            onRowToggle={(e: any) => this.setState({ expandedRows: e.data })}
            rowExpansionTemplate={this.rowExpansionTemplate}
          >
            <Column expander={true} style={{ width: "2em" }} />
            <Column field="Project" header="Name" />
            <Column field="Body" header="Description" />
            <Column field="Owner" header="Owner" body={this.ownerTemplate} />
            <Column field="Priority" header="Priority" />
            <Column field="MildStone" header="Mildstone" />
            <Column
              field="Status"
              header="Status"
              body={this.statusTemplate}
              style={{ padding: 0 }}
            />
          </DataTable>
        </div>
      </div>
    );
  }

  /* Api Call*/

  private getProjectList(): void {
    sp.web.lists
      .getByTitle("Project")
      .items.select(
        "Project_x0020_ID",
        "Project",
        "StartDate",
        "DueDate",
        "AssignedTo/Title",
        "AssignedTo/ID",
        "AssignedTo/EMail",
        "Status0/ID",
        "Status0/Status",
        "Status0/Status_x0020_Color",
        "Priority",
        "Body"
      )
      .expand("AssignedTo", "Status0")
      .getAll()
      .then((response: Array<CEOProjects>) => {
        this.getMildStones(response);
      });
  }

  private getMildStones(projectList: Array<CEOProjects>): void {
    sp.web.lists
      .getByTitle("Tasks List")
      .items.select("Title", "Project/ID", "Project/Title")
      .expand("Project")
      .filter("Duration eq 0")
      .get()
      .then((milestones: any[]) => {
        let timeline = new Array<ProjectTimeLine>();
        projectList.forEach(item => {
          let filteredMilestones = filter(milestones, function(milstoneItem) {
            return milstoneItem.Project.ID.toString() == item.Project_x0020_ID;
          });
          let mildstone = "";
          for (let count = 0; count < filteredMilestones.length; count++) {
            if (filteredMilestones.length == 0) {
              mildstone = filteredMilestones[count].Title;
            } else {
              if (count != filteredMilestones.length - 1) {
                mildstone = mildstone + filteredMilestones[count].Title + ",";
              } else {
                mildstone = mildstone + filteredMilestones[count].Title;
              }
            }
          }
          item.MildStone = mildstone;
          item.AssignedTo.forEach(element => {
            if (element.EMail != null) {
              element.imgURL =
                "https://esplrms-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/" +
                element.EMail.split("@")[0].toLowerCase() +
                "_esplrms_onmicrosoft_com_MThumb.jpg";
            } else {
              element.imgURL =
                "https://esplrms.sharepoint.com/sites/projects/SiteAssets/default.jpg";
            }
          });

          // Time Line
          timeline.push({
            id: item.Project_x0020_ID,
            name: item.Project,
            start: item.StartDate,
            end: item.DueDate,
            progress: 10
          });
        });
        this.setState({ projectList: projectList, projectTimeLine: timeline });
      });
  }

  private getTeamMembersByProject(id): void {
    
  }
  private getMildStonesByProject(id): void {
    sp.web.lists
      .getByTitle("Tasks List")
      .items.select("Title")
      .filter("Project/Title eq 'AlphaServe' and Duration eq 0")
      .get()
      .then((response: any[]) => {
        console.log("by pro -", response);
      });
  }

  private getKeyDocumentsByProject(id): void {
    sp.web.lists
      .getByTitle("Project Documents")
      .items.select("File", "Project/ID", "Project/Title")
      .expand("File", "Project")
      .filter("Project/Title eq 'AlphaServe'")
      .get()
      .then(response => {
        console.log("Project Documents -", response);
      });
  }

  private getTaggingByProject(id): void {
    sp.web.lists
      .getByTitle("Project Tags")
      .items.filter("Project/Title eq 'AlphaServe'")
      .get()
      .then(response => {
        console.log("Project Tags -", response);
      });
  }
}
