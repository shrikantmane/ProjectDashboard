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
import AddTask from '../AddTask/AddTask';
export default class ProjectTaskList extends React.Component<
  IProjectTaskListProps,
  IProjectTaskListState
  > {
  constructor(props) {
    super(props);
    this.state = {
      taskList: new Array<Task>(),
      showComponent: false,
      taskID: 0,
    };
    this.onAddTask = this.onAddTask.bind(this);
    this.reopenPanel = this.reopenPanel.bind(this);
    this.onEditTask = this.onEditTask.bind(this);
    this.deleteTask = this.deleteTask.bind(this);
    this.refreshGrid = this.refreshGrid.bind(this);
  }
  componentDidMount() {
    this.getAllTask();
  }
  refreshGrid() {
    this.setState({
      showComponent: false,
      taskID: null
    })
    this.getAllTask()
  }
  private getAllTask() {
    let filter = "Department/Department eq '" + this.props.department + "'";
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
            item.OwnerName = item.AssignedTo[0].Title;
          }
          item.Status = item.Status0 != null ? item.Status0.Status : "";
        });
        this.setState({ taskList: sortedResponse });
      });
  }
  reopenPanel() {
    this.setState({
      showComponent: false,
      taskID: null
    })
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

  startDateFilter(value, filter) {
    if (value && value != "" && filter && filter != "") {
      let startDate = moment(value).format("DD MMM YYYY");
      return startDate.toLowerCase().includes(filter.toLowerCase());
    } else {
      return true;
    }
  }

  endDateFilter(value, filter) {
    if (value && value != "" && filter && filter != "") {
      let endDate = moment(value).format("DD MMM YYYY");
      return endDate.toLowerCase().includes(filter.toLowerCase());
    } else {
      return true;
    }
  }
  onAddTask() {
    console.log('button clicked');
    this.setState({
      showComponent: true,
    });
  }
  onEditTask() {
    this.setState({
      showComponent: true,
      taskID: 1
    });
  }
  private deleteTask(id): void {
    sp.web.lists
      .getByTitle("All Tasks").items.getById(id).delete()
      .then(res => {
        console.log("res -", res);
        this.getAllTask();
      });
  }
  public render(): React.ReactElement<IProjectTaskListProps> {
    return (
      <div className="well recommendedProjects  ">
        <div className="row">
          <div className="col-sm-12 col-12 cardHeading">
            <div className="tasklist-div">
              <h5>Task List</h5>
              <button type="button" className="btn btn-primary btn-sm" style={{ marginBottom: "10px" }} onClick={this.onEditTask}>
                Edit Task
              </button>
              <button type="button" className="btn btn-primary btn-sm" style={{ marginBottom: "10px" }} onClick={this.deleteTask.bind(11)}>
                Delete Task
              </button>
              <button type="button" className="btn btn-primary btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddTask}>
                Add Task
              </button>
            </div>
          </div>

          <div className="clearfix" />
          {this.state.showComponent ?
            <AddTask id={this.state.taskID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} /> :
            null
          }
          <div className="col-sm-12 col-12 profileDetails-container" style={{ Width: "90%", marginLeft: "35px;" }}>
            <div>
              <DataTable value={this.state.taskList} rowGroupMode="subheader" groupField="Week" sortField="sort" sortOrder={1} scrollable={true} scrollHeight="200px"
                rowGroupHeaderTemplate={this.headerTemplate} rowGroupFooterTemplate={() => { return; }}>
                <Column field="Title" header="Title" filter={true} />
                <Column field="OwnerName" header="Owner" filter={true} />
                <Column field="StartDate" header="Start Date" sortable={true} body={this.startDateTemplate} filter={true} filterMatchMode="custom" filterFunction={this.startDateFilter} />
                <Column field="DueDate" header="Due Date" sortable={true} body={this.endDateTemplate} filter={true} filterMatchMode="custom" filterFunction={this.endDateFilter} />
                <Column field="Status" header="Status" body={this.statusTemplate} filter={true} />
              </DataTable>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
