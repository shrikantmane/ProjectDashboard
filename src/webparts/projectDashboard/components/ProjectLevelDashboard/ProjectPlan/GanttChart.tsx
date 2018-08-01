/*global gantt*/
import * as React from 'react';
// import 'dhtmlx-gantt';
import 'dhtmlx-gantt';
import 'dhtmlx-gantt/codebase/dhtmlxgantt.css';
declare var gantt: any;

export default class Gantt extends React.Component<any, any>{
  ganttContainer: any;

  constructor(props) {
    super(props);
    console.log('constructor');
    this.initGanttChart();
  }
  setZoom(value) {
    switch (value) {
      case 'Hours':
        gantt.config.scale_unit = 'day';
        gantt.config.date_scale = '%d %M';

        gantt.config.scale_height = 60;
        gantt.config.min_column_width = 30;
        gantt.config.subscales = [
          { unit: 'hour', step: 1, date: '%H' }
        ];
        break;
      case 'Days':
        gantt.config.min_column_width = 70;
        gantt.config.scale_unit = "week";
        gantt.config.date_scale = "#%W";
        gantt.config.subscales = [
          { unit: "day", step: 1, date: "%d %M" }
        ];
        gantt.config.scale_height = 60;
        break;
      case 'Months':
        gantt.config.min_column_width = 70;
        gantt.config.scale_unit = "month";
        gantt.config.date_scale = "%F";
        gantt.config.scale_height = 60;
        gantt.config.subscales = [
          { unit: "week", step: 1, date: "#%W" }
        ];
        break;
      default:
        break;
    }
  }

  shouldComponentUpdate(nextProps) {
    return this.props.zoom !== nextProps.zoom;
  }

  componentDidUpdate() {
    console.log('componentDidUpdate');
    gantt.render();
  }

  // initGanttEvents() {
  //   if(gantt.ganttEventsInitialized){
  //     return;
  //   }
  //   gantt.ganttEventsInitialized = true;

  //   gantt.attachEvent('onAfterTaskAdd', (id, task) => {
  //     if(this.props.onTaskUpdated) {
  //       this.props.onTaskUpdated(id, 'inserted', task);
  //     }
  //   });

  //   gantt.attachEvent('onAfterTaskUpdate', (id, task) => {
  //     if(this.props.onTaskUpdated) {
  //       this.props.onTaskUpdated(id, 'updated', task);
  //     }
  //   });

  //   gantt.attachEvent('onAfterTaskDelete', (id) => {
  //     if(this.props.onTaskUpdated) {
  //       this.props.onTaskUpdated(id, 'deleted');
  //     }
  //   });

  //   gantt.attachEvent('onAfterLinkAdd', (id, link) => {
  //     if(this.props.onLinkUpdated) {
  //       this.props.onLinkUpdated(id, 'inserted', link);
  //     }
  //   });

  //   gantt.attachEvent('onAfterLinkUpdate', (id, link) => {
  //     if(this.props.onLinkUpdated) {
  //       this.props.onLinkUpdated(id, 'updated', link);
  //     }
  //   });

  //   gantt.attachEvent('onAfterLinkDelete', (id, link) => {
  //     if(this.props.onLinkUpdated) {
  //       this.props.onLinkUpdated(id, 'deleted');
  //     }
  //   });
  // }

  componentDidMount() {
    console.log('componentDidMount');
    gantt.init(this.ganttContainer);
    gantt.parse(this.props.tasks);
    this.initGanttChart()
  }

  initGanttChart() {
    gantt.config.columns = [
      { name: "text", label: "Task name", tree: true, width: 80 },
      { name: "start_date", label: "Start time", align: "center", width: 80 },
      {
        name: "attachment1", label: "Attachment", width: 80,
        template: function (obj) {
          return "<a href='https://www.w3schools.com'>" + obj.attachment + "</a>";
        }
      },
      {
        name: "status", label: "Status", width: 80,
        template: function (obj) {
          return ("<span style='background:#ccc;'>" + obj.status + "</span>")
        }
      },
      {
        name: "duration", label: "Duration", align: "center", width: 80,
        template: function (obj) {
          console.log('duration', obj.duration);
          return ("<span style='background:#ccc;'>" + obj.duration + "</span>")
        }
      },
    ];
    gantt.config.layout = {
      css: "gantt_container",
      cols: [
        {
          width: 300,
          min_width: 200,

          // adding horizontal scrollbar to the grid via the scrollX attribute
          rows: [
            { view: "grid", scrollX: "gridScroll", scrollable: true, scrollY: "scrollVer" },
            { view: "scrollbar", id: "gridScroll" }
          ]
        },
        { resizer: true, width: 1 },
        {
          rows: [
            { view: "timeline", scrollX: "scrollHor", scrollY: "scrollVer" },
            { view: "scrollbar", id: "scrollHor" }
          ]
        },
        { view: "scrollbar", id: "scrollVer" }
      ]
    };
    //   gantt.templates.task_class  = function(start, end, task){
    //     switch (task.duration){
    //         case 1:
    //             return "test";
    //         case 2:
    //             return "test1";

    //     }
    // };
  }

  render() {
    console.log('render');
    this.setZoom(this.props.zoom);

    return (
      <div
        ref={(input) => { this.ganttContainer = input }}
        style={{ width: '100%', height: '100%' }}
      ></div>
    );
  }
}