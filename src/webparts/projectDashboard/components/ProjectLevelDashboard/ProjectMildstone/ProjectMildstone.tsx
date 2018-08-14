import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import moment from 'moment/src/moment';
import { ProgressBar } from 'react-bootstrap';
import { IProjectMildstoneProps } from './IProjectMildstoneProps';
import { IProjectMildstoneState } from './IProjectMildstoneState';
import { Milestone } from '../Project';
import styles from "../../ProjectDashboard.module.scss";
let progress:number;
export default class ProjectMildstone extends React.Component<IProjectMildstoneProps, IProjectMildstoneState> {

  constructor(props) {
    super(props);
    this.state = {
      milstoneList: new Array<Milestone>(),
      progressbar:null
    };
  }

  componentWillReceiveProps(nextProps) {
    if (this.props.scheduleList != nextProps.scheduleList)
      this.getAllMilestones(nextProps.scheduleList);
  }

  componentDidUpdate()
  {
     let container_width = 145 * document.getElementsByClassName("milestonesClass").length;
     let elements =  document.getElementsByClassName("row container-inner");
     if(elements && elements[0] && container_width > 0){
      elements[0].setAttribute("style", "width:" +  container_width.toString() + "px;");
     }
  }

  private getAllMilestones(scheduleList: string) {
    let temp1:any;
    let temp2:any;
    let temp3:number;
    let temp4:number;
    sp.web.lists.getByTitle(scheduleList).items
      .select('Title', 'DueDate', 'Status0/ID', 'Status0/Status', 'Status0/Status_x0020_Color', 'Priority').expand('Status0')
      .filter("Duration eq '0 days'")
      .get()
      .then((mildstones: Array<Milestone>) => {
        temp1=mildstones;
        this.setState({ milstoneList: mildstones })
       
    temp2=temp1.filter(item => item.Status0.Status == "Done" );
    temp3=temp1.length;
    temp4=temp2.length;
      console.log("a",temp1);
      console.log("b2",temp2)
      progress=((temp4/temp3)*100);
        console.log("milestones",mildstones)
        console.log("progress1",progress);
        this.setState({ progressbar: progress })
      });
  }

  public render(): React.ReactElement<IProjectMildstoneProps> {
    return (
      <div className="projectHealth">
        <div className="row">
          <div className="col-sm-3 col-12">
            <div className="row dark-blue">

              <div className="status-title">
                <span className="title-bullet inprogressStatus pull-left"></span>Health <span className="activityStatus"></span>
              </div>

              <div id="skill" className="mid-text mid-text col-sm-12 col-12">
                <h5>{this.state.progressbar}% Of Project is Done</h5>
                {/* <div className="progress-bar-outline"><span className="bar jquery"></span><h5>In Progress</h5></div> */}
                <ProgressBar bsClass="bar jquery" bsStyle="warning" now={this.state.progressbar} label="" />
                <div><h5>In Progress</h5></div>
              </div>
            </div>
          </div>
          <div className="col-sm-9 col-12 all-milestones container-outer">
            <div className="row container-inner">
              {this.state.milstoneList != null
                ? this.state.milstoneList.map((item, key) => {
                  return (
                    <div className="col-sm-4 col-6 milestonesClass">
                      <div className="milestoneList">
                        <h4 className="milestones-title" title={item.Title}><span className="title-bullet inprogressStatus pull-left" style={{ backgroundColor: item.Status0 ? item.Status0.Status_x0020_Color : "" }}></span>{item.Title}</h4>
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
                                <td>Status</td>
                                <td>{item.Status0 ? item.Status0.Status : ""}</td>
                              </tr>
                              <tr>
                                <td><i className="far fa-calendar-check"></i> Due Date</td>
                                <td>{moment(item.DueDate).format("DD MMM YYYY")}</td>
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
