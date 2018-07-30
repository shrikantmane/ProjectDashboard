import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { ProgressBar } from 'react-bootstrap';
import { IProjectMildstoneProps } from './IProjectMildstoneProps';
import { IProjectMildstoneState } from './IProjectMildstoneState';
import { Mildstone } from '../Project';
import styles from "../../ProjectDashboard.module.scss";

export default class ProjectMildstone extends React.Component<IProjectMildstoneProps, IProjectMildstoneState> {

  constructor(props) {
    super(props);
    this.state = {
      milstoneList: new Array<Mildstone>()
    };
  }
  // componentDidMount() {
  //   if(this.props.scheduleList)
  //   this.getAllMildstones(this.props.scheduleList);
  // }

  componentWillReceiveProps(nextProps) {
    if (this.props.scheduleList != nextProps.scheduleList)
      this.getAllMildstones(nextProps.scheduleList);
  }

  private getAllMildstones(scheduleList: string) {
    sp.web.lists.getByTitle(scheduleList).items
      .select('Title', 'DueDate', 'Status0/ID', 'Status0/Status', 'Status0/Status_x0020_Color', 'Priority').expand('Status0')
      .filter("Duration eq 0")
      .get()
      .then((mildstones: Array<Mildstone>) => {
        console.log("mildstones -", mildstones);
        this.setState({ milstoneList: mildstones })
      });
  }

  public render(): React.ReactElement<IProjectMildstoneProps> {
    return (
      <div className="projectHealth">
        <div className="row">
          <div className="col-sm-4 col-12">
            <div className="row dark-blue">

              <div className="status-title">
                <span className="title-bullet inprogressStatus pull-left"></span>Health <span className="activityStatus"><button type="button" className="btn-outline btn btn-sm">Project Outline</button></span>
              </div>

              <div id="skill" className="mid-text mid-text col-sm-12 col-12">
                <h5>80% Of Project is Done</h5>
                {/* <div className="progress-bar-outline"><span className="bar jquery"></span><h5>In Progress</h5></div> */}
                <ProgressBar bsClass="bar jquery" bsStyle="warning" now={80} label="" />
              <div><h5>In Progress</h5></div>
              </div>
            </div>
          </div>
          <div className="col-sm-8 col-12 all-milestones">
            <div className="row">
              {this.state.milstoneList != null
                ? this.state.milstoneList.map((item, key) => {
                  return (
                    <div className="col-sm-4 col-6">
                    <div className="milestoneList">
                      <h4 className="milestones-title"><span className="title-bullet inprogressStatus pull-left" style={{ backgroundColor: item.Status0 ? item.Status0.Status_x0020_Color : "" }}></span>{item.Title}</h4>
                      <div className="table-responsive-sm">
                        <table className="table">
                          <thead>
                            <tr>
                              <td></td>
                              <td></td>
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td><i className="far fa-calendar-check"></i></td>
                              <td>{item.DueDate}</td>
                            </tr>
                            <tr>
                              <td>Status</td>
                              <td>{item.Status0 ? item.Status0.Status : ""}</td>
                            </tr>
                            <tr>
                              <td>Stage</td>
                              <td>Devlopment</td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                      </div>
                    </div>
                  );
                })
                : null}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
