import * as React from 'react';
import * as ReactDom from 'react-dom';
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
    this.onScroll = this.onScroll.bind(this);
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
    document.getElementById('timelineMainDiv').scrollLeft = 900;   
    const node = ReactDom.findDOMNode(this);
    const mainElement = node.querySelector('.timelineMainDiv');
    mainElement.addEventListener('scroll', this.onScroll);
  }

  componentDidUpdate(){
    this.renderFrappeGanttDOM();
  }


  onScroll({ currentTarget }) {
    const node = ReactDom.findDOMNode(this);
    const asideElement = node.querySelector('.timelineProjectName');
    const mainElement = node.querySelector('.timelineMainDiv');   
    asideElement.scrollTop = mainElement.scrollTop;
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
    document.getElementById('timelineMainDiv').scrollLeft = 900;
  }

  private onWeekViewClick(){
    this.gantt.change_view_mode("Week");
    document.getElementById('timelineMainDiv').scrollLeft = 350;
  }

  private onMonthsViewClick(){
    this.gantt.change_view_mode("Month");
    document.getElementById('timelineMainDiv').scrollLeft = 250;
  }

  public render(): React.ReactElement<ICEOProjectTimeLineProps> {
    return (
       <div className="TimeLineContainer">
        <div className="timeLineBtnDiv">
          <button type="button" className="btn btn-default btn-sm timeLineBtn" onClick={this.onDayViewClick}>Day </button>
          <button type="button" className="btn btn-default btn-sm timeLineBtn" onClick={this.onWeekViewClick}>Week</button>
          <button type="button" className="btn btn-default btn-sm timeLineBtn" onClick={this.onMonthsViewClick}>Month</button>
        </div> 
          <div className="timelineMainDiv" id="timelineMainDiv">
            <div className="row">
              <div className="projectSideBar" id="projectSideBar">
              {/* <div className="timelineProjectNameDiv"> Project Name </div> */}
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
