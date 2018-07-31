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
    { id: 2, text: 'Task #2', start_date: '25-04-2017', attachment: 'attachment2', status: 'Completed', duration: 0, progress: 0, parent: 1 }
  ],
  links: [
    { id: 1, source: 1, target: 2, type: '0' }
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
        "Predecessors/Id"
      )
      .expand("Project", "Status0", "AssignedTo", "ParentID", "Predecessors")
      .get()
      .then((response: Array<Plan>) => {
        console.log("Plan", response);
        let chartData = new Chart();
        chartData.data = Array<ChartData>();
        chartData.link = Array<ChartLink>();

        response.forEach(item => {
          chartData.data.push({
            id: item.ID,
            text: item.Title,
            start_date: moment(item.StartDate).format("DD-MM-YYYY"),
            attachment: "",
            status: item.Status0 ? item.Status0.Status : "",
            duration: 3
            //duration: item.ID % 2 == 0 ? 1 : 2,
            // color : item.ID % 2 == 0 ? "red": "blue"
            //  parent : item.ID == 2 || item.ID == 3 ? 1 : null
          })
        });
        let rec = moment(chartData.data[0].start_date);

        this.setState({ chart: chartData });
      });
  }

  public render(): React.ReactElement<IProjectPlanProps> {
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
