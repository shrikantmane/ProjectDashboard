import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import moment from 'moment/src/moment';
import Gantt from './GanttChart';
import Toolbar from './Toolbar';
import { Plan, Chart, ChartData, ChartLink, Status } from "./Project";
import { IProjectPlanProps } from './IProjectPlanProps';
import { IProjectPlanState } from './IProjectPlanState';

export default class ProjectPlan extends React.Component<IProjectPlanProps, IProjectPlanState> {

  constructor(props) {
    super(props);
    this.state = {
      currentZoom: 'Months',
      chart: new Chart(),
      statusList: [],
      teamMembers : []
    };
    this.handleZoomChange = this.handleZoomChange.bind(this);
  }

  // componentDidMount(){
  //   let chartData = new Chart();
  //   chartData.data = new Array<ChartData>();
  //   chartData.links = new Array<ChartLink>(); 
  //   this.setState({ chart: chartData  });
  // }
  componentDidMount(){
    this.getTaskStatus();
  }
  componentWillReceiveProps(nextProps) {
    if (this.props.scheduleList != nextProps.scheduleList)
      this.getProjectPlan(nextProps.scheduleList);
      if (this.props.teamMemberlist != nextProps.teamMemberlist)
      this.getTeamMembersByProject(nextProps.teamMemberlist);
  }

  handleZoomChange(zoom) {
    this.setState({
      currentZoom: zoom
    });
  }

  private getTaskStatus(): void {
    sp.web.lists
        .getByTitle("Task Status Color").items
        .select("ID","Status","Status_x0020_Color","Sequence")
        .get()
        .then((response: Array<Status>) => {
          let statusList = [];
          console.log('response', response);
          response.forEach(element => {
            statusList.push({
              key: element.Status,
              label: element.Status,
              id: element.ID,
              color : element.Status_x0020_Color
            })
          });
          this.setState({ statusList : statusList });       
        })
}

 private getTeamMembersByProject(teamMemberlist: string): void {
        sp.web.lists
            .getByTitle(teamMemberlist)
            .items.select(
            "Team_x0020_Member/ID",
            "Team_x0020_Member/Title",
            "Team_x0020_Member/EMail",
            "Start_x0020_Date",
            "End_x0020_Date",
            "Status",
        )
            .expand("Team_x0020_Member")
            .get()
            .then((response: Array<any>) => {
              let teamMembers = [];
              response.forEach(element => {
                teamMembers.push({
                  key :element.Team_x0020_Member.Title,
                  label :element.Team_x0020_Member.Title,
                  id:element.Team_x0020_Member.ID,
                })
               });
               this.setState({ teamMembers : teamMembers });
            });
    }


  private getProjectPlan(scheduleList: string, ): void {
    sp.web.lists
      .getByTitle(scheduleList)
      .items.select(
        "ID",
        "Title",
        "StartDate",
        "DueDate",
        "Duration",
        "PercentComplete",
        "Body",
        "Status0/ID",
        "Status0/Status",
        "Status0/Status_x0020_Color",
        "AssignedTo/Title",
        "AssignedTo/ID",
        "AssignedTo/EMail",
        "Priority",
        "ParentID/Id",
        "Predecessors/Id",
        "Predecessors/Title"
      )
      .expand("Status0", "AssignedTo", "ParentID", "Predecessors")
      .get()
      .then((response: Array<Plan>) => {
        let chartData = new Chart();
        chartData.data = new Array<ChartData>();
        chartData.links = new Array<ChartLink>();
        let linkId = 0;
        response.forEach(item => {
          let duration = item.Duration && item.Duration != "" ? parseFloat(item.Duration.split(" ")[0]) : -1;
          chartData.data.push({
            id: item.ID,
            text: item.Title,
            body: item.Body,
            start_date: moment(item.StartDate).format("DD-MM-YYYY"),
            attachment: '',
            status: item.Status0 ?item.Status0.Status : "",
            priority :item.Priority,
            type: duration == 0 ? "Task2" : "Task" ,
            duration: duration,
            actualDuration: duration,
            progress: item.PercentComplete / 100,
            parent: item.ParentID ? item.ParentID.Id : null,
            comments: "",
            owner: item.AssignedTo.length > 0 ? item.AssignedTo[0].Title : "",          
            statusBackgroudColor: item.Status0 ? item.Status0.Status_x0020_Color : "",
          });
          item.Predecessors.forEach(predecessor => {
            chartData.links.push({
              id: linkId,
              source: predecessor.Id,
              target: item.ID,
              type: '0'
            })
          })
          linkId++;
        });
        console.log('chartData11111111111111', chartData);
        this.setState({ chart: chartData });
      });
  }

  public render(): React.ReactElement<IProjectPlanProps> {
    console.log('this.state.teamMembers', this.state.teamMembers);
    return (
      <div className="col-lg-12 col-md-12 col-sm-12 cardPadding">
        <div className="well recommendedProjects" style={{ maxHeight: 'none' }}>
          <div className="row">
            <div className="col-sm-12 col-12 cardHeading">
              <div className="tasklist-div">
                <h5>Project Plan</h5>
              </div>
            </div>
            <div className="col-sm-12 col-12">
              {this.state.chart && this.state.chart.data && this.state.statusList && this.state.statusList.length > 0 
              && this.state.teamMembers && this.state.teamMembers.length > 0 ?
                <div className="taskGanttContainer">
                  <Toolbar
                    zoom={this.state.currentZoom}
                    onZoomChange={this.handleZoomChange.bind(this)}
                  />
                  <div className="gantt-container">
                    <Gantt
                      scheduleList = {this.props.scheduleList}
                      tasks={this.state.chart}
                      statusList={this.state.statusList}
                      teamMembers={this.state.teamMembers}
                      zoom={this.state.currentZoom}
                    />
                  </div>
                </div>
                : null
              }
            </div>
          </div>
        </div>
      </div>
    );
  }
}
