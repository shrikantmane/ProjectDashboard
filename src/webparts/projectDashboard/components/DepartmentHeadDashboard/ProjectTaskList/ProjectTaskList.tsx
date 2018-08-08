import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import { IProjectTaskListProps } from "./IProjectTaskListProps";
import { IProjectTaskListState } from "./IProjectTaskListState";
import { Task, Week } from "../DepartmentDashboard";
import ReactTable from "react-table";
import "react-table/react-table.css";
import { sortBy } from "lodash";
import moment from "moment/src/moment";
export default class ProjectTaskList extends React.Component<
  IProjectTaskListProps,
  IProjectTaskListState
  > {
  constructor(props) {
    super(props);
    this.state = {
      taskList: new Array<Task>()
    };
  }

  componentDidMount() {
    this.getAllTask();
  }

  private getAllTask() {
    let filter = "Department/Department eq '"+ this.props.department + "'";
    sp.web.lists
      .getByTitle("All Tasks")
      .items.select(
        "Title",
        "StartDate",
        "DueDate",
        "Status0/ID",
        "Status0/Status",
        "Status0/Status_x0020_Color",
        "AssignedTo/ID",
        "AssignedTo/Title",
        "AssignedTo/EMail"
      )
      .expand("Status0", "AssignedTo")
      .filter(filter)
      .get()
      .then((response: Array<Task>) => {
        let sortedResponse = sortBy(response, function (dateObj) {
          return new Date(dateObj.StartDate);
        });
        sortedResponse.forEach(item => {
          let startOfWeek = moment().startOf("isoWeek");
          let endOfWeek = moment().endOf("isoWeek");
          let endOfNextWeek = moment(endOfWeek).add(7, "day");
          let currentDate = moment(item.StartDate);


          if (currentDate >= startOfWeek && currentDate <= endOfWeek) {
            item.Week = Week.CurrentWeek;
            item.Sort = 1;
          } else if (currentDate > endOfWeek && currentDate <= endOfNextWeek) {
            item.Week = Week.NextWeek;
            item.Sort = 2;
          } else if (currentDate > endOfNextWeek) {
            item.Week = Week.Future;
            item.Sort = 3;
          } else {
            item.Week = Week.Past;
            item.Sort = 4;
          }

          if (item.AssignedTo && item.AssignedTo.length > 0) {
            item.AssignedTo.forEach(element => {
              if (element.EMail != null) {
                element.ImgURL =
                  "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
                  element.EMail +
                  "&UA=0&size=HR64x64&sc=1531997060853";
              } else {
                element.ImgURL = "";
              }
            });
            item.OwnerName = item.AssignedTo[0].Title;          }
           item.Status = item.Status0 != null ? item.Status0.Status : "";
        });
          this.setState({ taskList: sortedResponse });
      });
  }

  headerTemplate(data) {
    return data.Week;
  }

  private statusTemplate(rowData: Task, column) {
    if (rowData && rowData.Status0 && rowData.Status0.Status != "") {
      return (
        <div>
          <div className="statusPill"
            style={{
              backgroundColor: rowData.Status0.Status_x0020_Color
            }}
          >
            {rowData.Status0.Status}
          </div>
        </div>
      )
    }
  }

  private startDateTemplate(rowData: Task, column) {
    return (
      <span>{moment(rowData.StartDate).format("DD MMM YYYY")}</span>
    );
  }

  private endDateTemplate(rowData: Task, column) {
    return (
    <span>{moment(rowData.DueDate).format("DD MMM YYYY")}</span>
    );
  }
  
  public render(): React.ReactElement<IProjectTaskListProps> {
    return (
      <div className="well recommendedProjects  ">
        <div className="row">
          <div className="col-sm-12 col-12 cardHeading">
            <div className="tasklist-div">
              <h5>Task List</h5>
            </div>
          </div>

          <div className="clearfix" />
          <div className="col-sm-12 col-12 profileDetails-container" style={{ Width: "90%", marginLeft: "35px;" }}>
            <div>
              <DataTable value={this.state.taskList} rowGroupMode="subheader" groupField="Week" sortField="sort" sortOrder={1} scrollable={true} scrollHeight="200px"
                rowGroupHeaderTemplate={this.headerTemplate} rowGroupFooterTemplate={() => { return; }}>
                <Column field="Title" header="Title" filter={true}/>
                <Column field="OwnerName" header="Owner"  filter={true}/>
                <Column field="StartDate" header="Start Date" sortable={true} body={this.startDateTemplate} filter={true}/>
                <Column field="DueDate" header="Due Date" sortable={true} body={this.endDateTemplate} filter={true}/>
                <Column field="Status" header="Status" body={this.statusTemplate} filter={true}/>
              </DataTable>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
