import * as React from 'react';
import GanttJS from "frappe-gantt";
import styles from './CEOProjectTimeLine.module.scss';
import { ICEOProjectTimeLineProps } from './ICEOProjectTimeLineProps';
import { ICEOProjectTimeLineState } from './ICEOProjectTimeLineState';
import CEOProjectTable from '../CEOProjectsTable/CEOProjectTable';

export default class CEOProjectTimeLine extends React.Component<ICEOProjectTimeLineProps, ICEOProjectTimeLineState> {  
 
  constructor(props) {
    super(props);
    this.onDayViewClick = this.onDayViewClick.bind(this);
    this.onWeekViewClick = this.onWeekViewClick.bind(this);
    this.onMonthsViewClick = this.onMonthsViewClick.bind(this);  
  }

  private gantt: any;
  private target: HTMLDivElement;
  componentWillReceiveProps(nextProps) {
    if (this.gantt && this.props.viewMode !== nextProps.viewMode) {
        this.gantt.change_view_mode(this.props.viewMode);
    }
  }
  
  componentDidMount() {
    this.renderFrappeGanttDOM();
  }

  componentDidUpdate(){
    this.renderFrappeGanttDOM();
  }

  private renderFrappeGanttDOM(): void { 
    this.gantt = new GanttJS(this.target, this.props.tasks, {
      on_click: this.props.onClick,
      on_date_change: this.props.onDateChange,
      on_progress_change: this.props.onProgressChange,
      on_view_change: this.props.onViewChange,
      custom_popup_html: this.props.customPopupHtml,     
    });
    //this._gantt.change_view_mode(this.props.viewMode);
  }

  private onDayViewClick(){
    this.gantt.change_view_mode("Day");
  }

  private onWeekViewClick(){
    this.gantt.change_view_mode("Week");
  }

  private onMonthsViewClick(){
    this.gantt.change_view_mode("Month ");
  }

  public render(): React.ReactElement<ICEOProjectTimeLineProps> {
    return (
       <div className="TimeLineContainer">
        <div className="timeLineBtnDiv">
          <button type="button" className="btn btn-default btn-sm timeLineBtn" onClick={this.onDayViewClick}>Day </button>
          <button type="button" className="btn btn-default btn-sm timeLineBtn" onClick={this.onWeekViewClick}>Week</button>
          <button type="button" className="btn btn-default btn-sm timeLineBtn" onClick={this.onMonthsViewClick}>Month</button>
        </div> 
          <div className="timelineMainDiv">
            <div className="row">
              <div className="col-md-2 col-6">
              <div className="timelineProjectNameDiv"> Project Name </div>
                <div className="timelineProjectName">
                  { 
                        this.props.tasks.map(function(item, index){
                            return <div key={index} style={{height:"38px"}}>{item.name} </div>
                        }) 
                    }
                </div>
              </div>
              <div className="col-md-10 col-5">
                  <div ref={r => this.target = r} />
              </div>
            </div>
          </div>              
      </div>
    );
  }
}
