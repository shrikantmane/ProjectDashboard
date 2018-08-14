/*global gantt*/
import * as React from 'react';
// import 'dhtmlx-gantt';
import 'dhtmlx-gantt';
import 'dhtmlx-gantt/codebase/dhtmlxgantt.css';
import 'dhtmlx-gantt/codebase/ext/dhtmlxgantt_tooltip.js';
import { IGanttChartProps } from './IGanttChart';
import Documents from '../Documents/Documents';
import Comments from '../Comments/Comments';
declare var gantt: any;
var currentScope: any;
export default class Gantt extends React.Component<IGanttChartProps, {
  showDocumentComponent: boolean;
  showCommentComponent: boolean;
  taskID: any,
  documentID: any
}>{
  ganttContainer: any;

  constructor(props) {
    super(props);
    this.state = {
      showDocumentComponent: false,
      showCommentComponent: false,
      taskID: 0,
      documentID: 0
    }
    this.onDocuments = this.onDocuments.bind(this);
    this.onComment = this.onComment.bind(this);
    this.reopenPanel = this.reopenPanel.bind(this);
    this.initGanttChart();
    console.log('schedulelist : ', this.props.scheduleList);
    console.log('commentList : ', this.props.commentList);
  }
  onDocuments(id): void {
    this.props.showDocuments(Number(id));
  }
  onComment(id): void {
    this.props.showComments(Number(id));
  }
  reopenPanel() {
    this.setState({
      showDocumentComponent: false,
      showCommentComponent: false,
      taskID: null,
      documentID: null
    })
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
    currentScope = this;
    gantt.init(this.ganttContainer);
    gantt.clearAll();
    gantt.parse(this.props.tasks);
    this.initGanttChart()
  }

  initGanttChart() {
    gantt.config.columns = [
      {
        name: "attachment", label: "", width: 30,
        template: function (obj) {
          return "<i class='fas fa-paperclip'></i>";
        }
      },
      {
        name: "comments", label: "", width: 30,
        template: function (obj) {
          return "<i class='far fa-comments'></i>";
        }
      },
      { name: "text", label: "Task name", tree: true, width: 100 },
      { name: "start_date", label: "Start time", align: "center", width: 80 },
      {
        name: "ryg", label: "RYG", width: 22,
        template: function (obj) {
          return ("<div class='title-bullet' style='background-color:" + obj.statusBackgroudColor + "'></div>")
        }
      },
      {
        name: "status", label: "Status", width: 80,
        template: function (obj) {
          return (obj.status)
        }
      },
      {
        name: "actualDuration", label: "Duration", align: "center", width: 80,
        template: function (obj) {
          return ("<span>" + obj.actualDuration + "</span>")
        }
      },
    ];
    gantt.attachEvent("onTaskClick", function (id, e) {
      //any custom logic here
      console.log(e);
      console.log(id);
      if (e.target.className === 'fas fa-paperclip') {
        currentScope.onDocuments(id);
      }
      else if (e.target.className === 'far fa-comments') {
        currentScope.onComment(id);
      }
      return true;
    });
    gantt.templates.grid_file = function (item) {
      return "";
    };
    gantt.config.readonly = true;
    gantt.templates.tooltip_text = function (start, end, task) {
      let label = task.actualDuration == 0 ? "Milestone" : "Task"
      return "<div><b> " + label + "</b> " + task.text + "<br/><b>Start Date :</b> " + new Date(start).toDateString();
    };
    gantt.config.layout = {
      css: "gantt_container",
      cols: [
        {
          width: 330,
          min_width: 250,

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
    gantt.templates.task_class = function (start, end, task) {
      if (task.actualDuration == 0) {
        return "milestone";
      }
    };
  }

  render() {
    this.setZoom(this.props.zoom);
    return (
      <div
        ref={(input) => { this.ganttContainer = input }}
        style={{ width: '100%', height: '100%' }}
      >
      </div>
    );
  }
}