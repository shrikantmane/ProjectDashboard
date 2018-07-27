import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectTaskListProps } from './IProjectTaskListProps';
import { IProjectTaskListState } from './IProjectTaskListState';
import { Task } from '../Project';

export default class ProjectTaskList extends React.Component<IProjectTaskListProps, IProjectTaskListState> {

  constructor(props) {
    super(props);
    this.state = {
      taskList: new Array<Task>()
    };
  }

  componentWillReceiveProps(nextProps) {
    if (this.props.scheduleList != nextProps.scheduleList)
      this.getAllTask(nextProps.scheduleList);
  }

  private getAllTask(scheduleList: string) {
    sp.web.lists.getByTitle(scheduleList).items
      .select('Title', 'Status0/ID', 'Status0/Status', 'Status0/Status_x0020_Color', 'AssignedTo/ID', 'AssignedTo/Title').expand('Status0', 'AssignedTo')
      .get()
      .then((response:  Array<Task>) => {
        console.log("All Task -", response);
        this.setState({taskList : response});
      });
  }

  public render(): React.ReactElement<IProjectTaskListProps> {
    return (
      <div className="col-lg-6 col-md-6 col-sm-12">
        <div className="well recommendedProjects  ">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <div className="tasklist-div">
                <h5>Task List</h5>
                <span>
                  <a className="btn btn-primary btn-lg btn-md btn-sm btn-xs ">Create Task</a>
                </span>
              </div>
            </div>

            <div className="clearfix"></div>
            <div className="profileDetails-container">
              <div className="table-responsive">
                <table className="table table-striped ">
                </table>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
