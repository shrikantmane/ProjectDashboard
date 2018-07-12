import * as React from 'react';
import GanttJS from "frappe-gantt";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { ICEOProjectTimeLineProps } from './ICEOProjectTimeLineProps';
import { ICEOProjectTimeLineState } from './ICEOProjectTimeLineState';
import CEOProjectTable from '../CEOProjectsTable/CEOProjectTable'

export default class CEOProjectTimeLine extends React.Component<ICEOProjectTimeLineProps, ICEOProjectTimeLineState> {  
 
  constructor(props) {
    super(props);
  }

  private gantt: any;
  private target: HTMLDivElement;
  componentWillReceiveProps(nextProps) {
    if (this.gantt && this.props.viewMode !== nextProps.viewMode) {
        this.gantt.change_view_mode(this.props.viewMode);
    }
  }

  componentDidMount() {
    SPComponentLoader.loadCss("./TimeLine.module.scss");
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

  public render(): React.ReactElement<ICEOProjectTimeLineProps> {
    return (
           
      <div style={{ marginTop: "10px", height:"320px", overflow:"auto"}}>
        <div className="row">
          <div className="col-md-2">
            <div style={{ float:"left", marginTop:"60px", fontWeight:"bold", fontSize: "12px", color:"grey" }}>
              { 
                    this.props.tasks.map(function(item, index){
                        return <div key={index} style={{height:"38px", paddingTop:"9px"}}>{item.name} </div>
                    }) 
                }
            </div>
          </div>
          <div className="col-md-10">
              <div ref={r => this.target = r} />
          </div>
        </div>
      </div>
    );
  }
}
