import * as React from 'react';
import * as ReactDom from 'react-dom';
import { ITimeLineChartProps } from './ITimeLineChartProps';
import { ITimeLineChartState } from './ITimeLineChartState';
import Timeline from 'react-calendar-timeline/lib';
import 'react-calendar-timeline/lib/Timeline.css';
import moment from 'moment/src/moment';

export default class TimeLineChart extends React.Component<ITimeLineChartProps, ITimeLineChartState> {  
 
  constructor(props) {
    super(props);
    
  }

  public render(): React.ReactElement<ITimeLineChartProps> {
    return (
       <div className="react-calendar-timeline-div">
        <Timeline
        sidebarWidth={200}
        lineHeight ={35}
        canMove = {false}
        canResize = {false}
        canChangeGroup = {false}
        fixedHeader = 'none'
        stickyHeader ='sticky'
        groups={this.props.groups}
        items={this.props.items}
        defaultTimeStart={moment().add(-15, 'day')}
        defaultTimeEnd={moment().add(15, 'day')}
        sidebarContent= {<p>Projects</p>}
        groupRenderer = {({ group }) => {
          return (
            <span title={group.title}>{group.title}</span>
          )
        }
      }
      />                     
      </div>
    );
  }
}
