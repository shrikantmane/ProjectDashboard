import * as React from 'react';
import * as ReactDom from 'react-dom';
import GanttJS from "frappe-gantt";
import styles from './CEOProjectTimeLine.module.scss';
import { ICEOProjectTimeLineProps } from './ICEOProjectTimeLineProps';
import { ICEOProjectTimeLineState } from './ICEOProjectTimeLineState';
import Timeline from 'react-calendar-timeline/lib';
import 'react-calendar-timeline/lib/Timeline.css';
import moment from 'moment/src/moment';

export default class CEOProjectTimeLine extends React.Component<ICEOProjectTimeLineProps, ICEOProjectTimeLineState> {  
 
  constructor(props) {
    super(props);
    
  }

  public render(): React.ReactElement<ICEOProjectTimeLineProps> {
    return (
       <div>
        <Timeline
        sidebarWidth={200}
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
