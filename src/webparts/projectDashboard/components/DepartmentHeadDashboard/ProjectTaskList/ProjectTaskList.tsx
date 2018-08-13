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
import Documents from '../Documents/Documents';
import Comments from '../Comments/Comments';
export default class ProjectTaskList extends React.Component<
  IProjectTaskListProps,
  IProjectTaskListState
  > {
  constructor(props) {
    super(props);
    this.state = {
      taskList: new Array<Task>(),
      showComponent: false,
      showDocumentComponent: false,
      showCommentComponent: false,
      taskID: 0,
      documentID: 0
    };
    this.onAddTask = this.onAddTask.bind(this);
    this.reopenPanel = this.reopenPanel.bind(this);
    this.onDocuments = this.onDocuments.bind(this);
    this.onComment = this.onComment.bind(this);
    this.refreshGrid = this.refreshGrid.bind(this);
    this.editTemplate = this.editTemplate.bind(this);
    this.deleteTemplate = this.deleteTemplate.bind(this);
    this.deleteListItem = this.deleteListItem.bind(this);
    this.rowClassName = this.rowClassName.bind(this);
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
      "ID",
      "Title",
      "StartDate",
      "DueDate",
      "Status0/ID",
      "Status0/Status",
      "Status0/Status_x0020_Color",
      "AssignedTo/ID",
      "AssignedTo/Title",
      "AssignedTo/EMail",
      "IsRemoved"
      )
      .expand("Status0", "AssignedTo")
      .filter(filter)
      .get()
      .then((response: Array<Task>) => {
        console.log('Task List:', response);
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
      showDocumentComponent: false,
      taskID: null
    })
  }
  headerTemplate(data) {
    return data.Week;
  }
  editTemplate(rowData, column) {
    if (rowData.IsRemoved) {
      return <div style={{ display: "none" }}></div>
    } else {
      return <a href="#" onClick={this.onEditProject.bind(this, rowData)}><i className="far fa-edit"></i> </a>;
    }
  }
  deleteTemplate(rowData, column) {
    if (rowData.IsRemoved) {
      return <div style={{ display: "none" }}></div>
    } else {
      return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}><i className="fas fa-trash"></i></a>;
    }
  }
  private onEditProject(rowData, e): any {
    e.preventDefault();
    console.log('Edit :' + rowData);
    this.setState({
      showComponent: true,
      taskID: rowData.ID
    });
  }
  private deleteListItem(rowData, e): any {
    var result = confirm("Are you sure you want to remove this task?");
    if (result) {
      e.preventDefault();
      console.log('Edit :' + rowData.ID);
      sp.web.lists
        .getByTitle("All Tasks")
        .items.getById(rowData.ID).update({
          IsRemoved: true
        })
        .then(res => {
          console.log("res -", res);
          this.getAllTask();
        });
    }
  }
  rowClassName(rowData) {
    let removedClass = rowData.IsRemoved;
    return { 'ui-state-highlight': (removedClass === true) };
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
  private ownerTemplate(rowData: Task, column) {
    if (rowData.AssignedTo)
      return (
        <div className="ownerImage">
          <img src={rowData.AssignedTo[0].ImgURL} />
          <div className="ownerName">{rowData.AssignedTo[0].Title}</div>
        </div>
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
  onDocuments(id): void {
    this.setState({
      showDocumentComponent: true,
      documentID: 1
    });
  }
  onComment(id): void {
    this.setState({
      showCommentComponent: true,
      documentID: 1
    });
  }
  public render(): React.ReactElement<IProjectTaskListProps> {
    return (
      <div className="well recommendedProjects taskListContainer  ">
        <div className="row">
          <div className="col-sm-12 col-12 cardHeading">
            <div className="tasklist-div">
              <h5>Task List</h5>
              {/* <button type="button" className="btn btn-primary btn-sm" style={{ marginBottom: "10px" }} onClick={this.onDocuments.bind(11)}>
                Documents
              </button>
              <button type="button" className="btn btn-primary btn-sm" style={{ marginBottom: "10px" }} onClick={this.onComment.bind(11)}>
                Comments
              </button> */}
              <button type="button" className="btn btn-primary btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddTask}>
                Add Task
              </button>
            </div>
          </div>

          <div className="clearfix" />
          {this.state.showCommentComponent ?
            <Comments id={this.state.documentID} parentReopen={this.reopenPanel} /> :
            null
          }
          {this.state.showDocumentComponent ?
            <Documents id={this.state.documentID} parentReopen={this.reopenPanel} /> :
            null
          }
          {this.state.showComponent ?
            <AddTask id={this.state.taskID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} /> :
            null
          }
          <div className="col-sm-12 col-12 profileDetails-container" style={{ Width: "90%", marginLeft: "35px;" }}>
            <div>
              <DataTable value={this.state.taskList} rowGroupMode="subheader" groupField="Week" sortField="sort" sortOrder={1} scrollable={true} scrollHeight="200px"
                rowGroupHeaderTemplate={this.headerTemplate} rowGroupFooterTemplate={() => { return; }} rowClassName={this.rowClassName} responsive ={true} >
                <Column body={this.editTemplate} style={{ width: "2%"}} />
                <Column field="Title" header="Title" filter={true} style={{ width: "25%"}} />
                <Column field="OwnerName" header="Owner"  filter={true}  body={this.ownerTemplate}  style={{ width: "20%"}}/>
                <Column field="StartDate" header="Start Date" sortable={true} body={this.startDateTemplate} filter={true}  style={{ width: "15%"}} filterMatchMode="custom" filterFunction={this.startDateFilter} />
                <Column field="DueDate" header="Due Date" sortable={true} body={this.endDateTemplate} filter={true} filterMatchMode="custom"  style={{ width: "15%"}} filterFunction={this.endDateFilter} />
                <Column field="Status" header="Status" body={this.statusTemplate}  style={{ width: "15%"}} filter={true} />
                <Column body={this.deleteTemplate} style={{ width: "7%" }} />
              </DataTable>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
