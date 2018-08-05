import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectTaskListProps } from "./IProjectTaskListProps";
import { IProjectTaskListState } from "./IProjectTaskListState";
import { Task, Week } from "../Project";
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
      taskList: new Array<Task>(),
      expanded: { 0: true }
    };
  }

  componentWillReceiveProps(nextProps) {
    if (this.props.taskList != nextProps.taskList)
      this.getAllTask(nextProps.taskList);
  }

  private getAllTask(taskList: string) {
    sp.web.lists
      .getByTitle(taskList)
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
      .get()
      .then((response: Array<Task>) => {
        let sortedResponse = sortBy(response, function (dateObj) {
          return new Date(dateObj.StartDate);
        });
        let pastRecord = false;
        let currentRecord = false;
        sortedResponse.forEach(item => {
          let startOfWeek = moment().startOf("isoWeek");
          let endOfWeek = moment().endOf("isoWeek");
          let endOfNextWeek = moment(endOfWeek).add(7, "day");
          let currentDate = moment(item.StartDate);


          if (currentDate >= startOfWeek && currentDate <= endOfWeek) {
            item.Week = Week.CurrentWeek;
            currentRecord = true;
          } else if (currentDate > endOfWeek && currentDate <= endOfNextWeek) {
            item.Week = Week.NextWeek;
          } else if (currentDate > endOfNextWeek) {
            item.Week = Week.Future;
          } else {
            item.Week = Week.Past;
            pastRecord = true;
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
            item.OwnerName = item.AssignedTo[0].Title
          }
          item.Status = item.Status0 != null ? item.Status0.Status : "";
        });
        if (pastRecord && currentRecord) {
          this.setState({
            expanded: {
              0: false,
              1: true,
            }, 
            taskList: sortedResponse
          });
        } else {
          this.setState({ taskList: sortedResponse });
        }
      });
  }

  handleRowExpanded(rowsState, index) {
    this.setState({
      expanded: {
        [index[0]]: !this.state.expanded[index[0]],
      },
    });
  }

  public render(): React.ReactElement<IProjectTaskListProps> {
    return (
      <div className="col-md-7 col-xs-12 taskListPadding">
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
                <ReactTable
                  collapseOnSortingChange={false}
                  collapseOnDataChange={false}
                  showPagination={true}
                  data={this.state.taskList}
                  filterable
                  expanded={this.state.expanded}
                  defaultFilterMethod={(filter, row, column) => {
                    const id = filter.pivotId || filter.id; 
                    return row[id] !== undefined ? String(row[id]).toLocaleLowerCase().match(filter.value.toLocaleLowerCase()) : true
                  }}
                  onFilterChange ={()=>{
                    if((this.state.expanded[0]) == false ){
                      this.setState
                      ({ expanded: { 0: true }})
                    }
                  }}
                  onExpandedChange={(newExpanded, index, event) => this.handleRowExpanded(newExpanded, index)}
                  columns={[
                    {
                      Header: "",
                      accessor: "Week",
                      sortable: false,
                      filterable: false,
                    },
                    {
                      Header: "Title",
                      accessor: "Title",
                      Cell: ({ row, original }) => {
                        if (original && original.Title) {
                          return (
                            <div className="taskListTitle">
                              <span title={original.Title}> {original.Title}</span>
                            </div>
                          )
                        }
                      },
                      Aggregated: row => {
                        return (
                          <span></span>
                        );
                      }
                    },
                    {
                      Header: "Owner",
                      accessor: "OwnerName",
                      Cell: ({ row, original }) => {
                        if (original && original.AssignedTo && original.AssignedTo.length > 0) {
                          return (
                            <div>
                              {/* <img src={original.AssignedTo[0].ImgURL} className="img-responsive"></img> */}
                              <span> {original.AssignedTo[0].Title}</span>
                            </div>
                          )
                        }
                      },
                      Aggregated: row => {
                        return (
                          <span></span>
                        );
                      }
                    },
                    {
                      Header: "Status",
                      accessor: "Status",
                      Cell: ({ row, original }) => {
                        if (original && original.Status0 && original.Status0.Status != "") {
                          return (
                            <div>
                              <div className="statusPill"
                                style={{
                                  backgroundColor: original.Status0.Status_x0020_Color
                                }}
                              >
                                {original.Status0.Status}
                              </div>
                            </div>
                          )
                        }
                      },
                      Aggregated: row => {
                        return (
                          <span></span>
                        );
                      }
                    }
                  ]}
                  pivotBy={["Week"]}
                  defaultPageSize={4}
                  className="-striped -highlight"
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
