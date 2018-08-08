import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import moment from 'moment/src/moment';
import Gantt from './GanttChart';
import Toolbar from './Toolbar';
import { Plan, Chart, ChartData, ChartLink } from "../Project";
import { IProjectPlanProps } from './IProjectPlanProps';
import { IProjectPlanState } from './IProjectPlanState';

export default class ProjectPlan extends React.Component<IProjectPlanProps, IProjectPlanState> {

  constructor(props) {
    super(props);
    this.state = {
      currentZoom: 'Months',
       chart: new Chart(),
    };
    this.handleZoomChange = this.handleZoomChange.bind(this);
  }


  componentWillReceiveProps(nextProps) {
    if (this.props.scheduleList != nextProps.scheduleList)
      this.getProjectPlan(nextProps.scheduleList);
  }

  handleZoomChange(zoom) {
    this.setState({
      currentZoom: zoom
    });
  }

  private getProjectPlan(scheduleList: string): void {
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
        chartData.data = Array<ChartData>();
        chartData.links = Array<ChartLink>(); 
        let linkId = 0;
        response.forEach(item => {
          chartData.data.push({
            id: item.ID, 
            text: item.Title, 
            start_date:  moment(item.StartDate).format("DD-MM-YYYY"), 
            attachment: '', 
            status: item.Status0 ? item.Status0.Status : "", 
            duration: item.Duration && item.Duration != "" ? parseFloat(item.Duration.split(" ")[0]) : -1,
            actualDuration :item.Duration && item.Duration != "" ? parseFloat(item.Duration.split(" ")[0]) : -1,
            progress: item.PercentComplete / 100,            
            parent:item.ParentID ? item.ParentID.Id :null,
            comments: "",
            statusBackgroudColor : item.Status0 ? item.Status0.Status_x0020_Color : "", 
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
        })      
        this.setState({ chart: chartData  });
      });
  }

  public render(): React.ReactElement<IProjectPlanProps> {
    return (
      <div className="col-lg-12 col-md-12 col-sm-12 cardPadding">
        <div className="well recommendedProjects" style={{maxHeight:'none'}}>
          <div className="row">
            <div className="col-sm-12 col-12 cardHeading">
              <div className="tasklist-div">
                <h5>Project Plan</h5>
              </div>
            </div>
            <div className="col-sm-12 col-12">
              {this.state.chart && this.state.chart.data && this.state.chart.data.length > 0 ?
                <div className="taskGanttContainer">
                  <Toolbar
                    zoom={this.state.currentZoom}
                    onZoomChange={this.handleZoomChange.bind(this)}
                  />
                  <div className="gantt-container">
                    <Gantt
                      tasks={this.state.chart}
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
