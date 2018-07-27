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
    if(this.props.scheduleList != nextProps.scheduleList)
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
      <div className="row">
        <div className="col-sm-4 col-md-3 col-lg-3">
          <div className="row dark-blue">

            <div className="status-title">
              <span className="title-bullet inprogressStatus pull-left"></span>Health <span className="pull-right activityStatus"><button type="button" className="btn-outline btn btn-sm">Project Outline</button></span>
            </div>

            <div id="skill" className="mid-text">
              <h5>80% Of Project is Done</h5>
              {/* <div className="progress-bar-outline"><span className="bar jquery"></span><h5>In Progress</h5></div> */}
              <ProgressBar bsStyle="warning" now={80} label={`${80}%`} />;
              <div><h5>In Progress</h5></div>
            </div>
          </div>
        </div>
        <div className="col-lg-9 col-md-9 col-sm-8 all-milestones">
          <div className="row">
            {this.state.milstoneList != null
              ? this.state.milstoneList.map((item, key) => {
                return (
                  <div className="col-lg-2 col-md-3 col-sm-4 col-xs-6">
                    <h4 className="milestones-title"><span className="title-bullet inprogressStatus pull-left" style={{backgroundColor : item.Status0 ? item.Status0.Status_x0020_Color : "" }}></span>{item.Title}</h4>
                    <table>
                      <thead>
                        <tr>
                          <td></td>
                          <td></td>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td><i className="far fa-calendar-check"></i></td>
                          <td>  {item.DueDate}</td>
                        </tr>
                        <tr>
                          <td>Status</td>
                          <td>  {item.Status0 ? item.Status0.Status : ""}</td>
                        </tr>
                        <tr>
                          <td>Stage</td>
                          <td>  Devlopment</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                );
              })
              : null}
          </div>
        </div>
      </div>
    );
  }
}
