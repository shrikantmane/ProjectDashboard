import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import moment from 'moment/src/moment';
import Gantt from './GanttChart';
import Toolbar from './Toolbar';
import { Plan, Chart, ChartData, ChartLink } from "../Project";
import { IProjectPlanProps } from './IProjectPlanProps';
import { IProjectPlanState } from './IProjectPlanState';
let data = {
  data: [
    { id: 1, text: 'Task #1', start_date: '15-04-2017', attachment: 'attachment1', status: 'In Progress', duration: 3, progress: 0.6 },
    { id: 2, text: 'Task #2', start_date: '25-04-2017', attachment: 'attachment2', status: 'Completed', duration: 2, progress: 0, parent: 1 },
    { id: 3, text: 'Task #3', start_date: '26-04-2017', attachment: 'attachment2', status: 'Completed', duration: 2, progress: 0 },
    { id: 4, text: 'Task #4', start_date: '03-05-2017', attachment: 'attachment2', status: 'Completed', duration: 2, progress: 0 }
  ],
  links: [
    { id: 1, source: 2, target: 1, type: '0' },
    { id: 2, source: 3, target: 4, type: '0' }
  ]
};

export default class ProjectPlan extends React.Component<IProjectPlanProps, IProjectPlanState> {

  constructor(props) {
    super(props);
    this.state = {
      currentZoom: 'Days',
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
        "Project/ID",
        "Project/Title",
        "AssignedTo/Title",
        "AssignedTo/ID",
        "AssignedTo/EMail",
        "Priority",
        "ParentID/Id",
        "Predecessors/Id",
        "Predecessors/Title"
      )
      .expand("Project", "Status0", "AssignedTo", "ParentID", "Predecessors")
      .get()
      .then((response: Array<Plan>) => {
        console.log("Plan", response);
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
            parent:item.ParentID ? item.ParentID.Id :null
          });
           item.Predecessors.forEach(predecessor => {
            chartData.links.push({
              id: linkId, 
              source: item.ID, 
              target:  predecessor.Id , 
              type: '0'
            })
           })
           linkId++;
        })      
        this.setState({ chart: chartData  });
      });
  }

  public render(): React.ReactElement<IProjectPlanProps> {
    console.log('this.state.chart', this.state.chart);
    return (
      <div>
        {this.state.chart && this.state.chart.data && this.state.chart.data.length > 0 ?
          <div>
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
    );
  }
}
