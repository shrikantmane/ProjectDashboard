/*global gantt*/
import * as React from 'react';
import { sp, ItemAddResult, EmailProperties } from "@pnp/sp";
import { find } from "lodash";
// import 'dhtmlx-gantt';
import 'dhtmlx-gantt';
import 'dhtmlx-gantt/codebase/dhtmlxgantt.css';
import 'dhtmlx-gantt/codebase/ext/dhtmlxgantt_tooltip.js';
import { IGanttChartProps } from './IGanttChart';
import { Chart, ChartData, ChartLink } from "./Project";
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
        gantt.config.date_scale = "Week %W";
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
          { unit: "week", step: 1, date: "week %W" }
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

  initGanttEvents() {
    if (gantt.ganttEventsInitialized) {
      return;
    }
    gantt.ganttEventsInitialized = true;

    gantt.attachEvent('onAfterTaskAdd', (id, task: any) => {
      let duration = task.type === "Task2" ? 0 : task.duration;
      let status: any = find(this.props.statusList, { key: task.status });
      let owner: any = find(this.props.teamMembers, { key: task.owner });

      const emailProps: EmailProperties = {
        To: [owner.email],
        Subject: "Task added",
        Body: `<b> New Task has been added.</b>`
      };

      sp.web.lists.getByTitle(this.props.scheduleList).items.add({
        Title: task.text,
        StartDate: task.start_date, //"2018-08-03T07:00:00Z"
        DueDate: task.end_date,   //"2018-08-03T07:00:00Z" 
        Status0Id: status.id,
        Body: task.body,
        AssignedToId: { results: [owner.id] },
        Duration: duration + " days",
        Priority: task.priority,
        // ParentId: task.parent         

      }).then((iar: ItemAddResult) => {
        sp.utility.sendEmail(emailProps).then(mail => {
        });
      }).catch(err => {
        console.log("Error while adding ScheduleTask", err);
      });
    });

    gantt.attachEvent('onAfterTaskUpdate', (id, task) => {
      let duration = task.type === "Task2" ? 0 : task.duration;
      let status: any = find(this.props.statusList, { key: task.status });
      let owner: any = find(this.props.teamMembers, { key: task.owner });
      const emailProps: EmailProperties = {
        To: [owner.email],
        Subject: "Task updated",
        Body: `<b> Task has been updated.</b>`
      };
      sp.web.lists.getByTitle(this.props.scheduleList).items.getById(task.id).update({
        Title: task.text,
        StartDate: task.start_date, //"2018-08-03T07:00:00Z"
        DueDate: task.end_date,   //"2018-08-03T07:00:00Z" 
        Status0Id: status.id,
        Body: task.body,
        AssignedToId: { results: [owner.id] },
        Duration: duration + " days",
        Priority: task.priority,
        //ParentId: task.parent
      }).then((iar: ItemAddResult) => {
        sp.utility.sendEmail(emailProps).then(mail => {
        });
      }).catch(err => {
        console.log("Error while updating ScheduleTask", err);
      });
    });

    gantt.attachEvent('onAfterTaskDelete', (id, task) => {
      sp.web.lists.getByTitle(this.props.scheduleList).items.getById(task.id).delete().then(_ => {
      }).catch(err => {
        console.log("Error while deleting ScheduleTask", err);
      });

    });

    gantt.attachEvent('onAfterLinkAdd', (id, link) => {
      if (this.props.onLinkUpdated) {
        this.props.onLinkUpdated(id, 'inserted', link);
      }
    });

    gantt.attachEvent('onAfterLinkUpdate', (id, link) => {
      if (this.props.onLinkUpdated) {
        this.props.onLinkUpdated(id, 'updated', link);
      }
    });

    gantt.attachEvent('onAfterLinkDelete', (id, link) => {
      if (this.props.onLinkUpdated) {
        this.props.onLinkUpdated(id, 'deleted');
      }
    });
    //   gantt.attachEvent("onTaskCreated", function(id, task){
    //     console.log('onTaskCreated',task);
    //     task.statusBackgroudColor = "blue";
    //     return true;
    // });
    gantt.attachEvent("onLightboxSave", function (id, item) {
      let status: any = find(currentScope.props.statusList, { key: item.status });
      item.statusBackgroudColor = status.color;
      item.actualDuration = item.type === "Task2" ? 0 : item.duration;
      item.duration = item.type === "Task2" ? 1 : item.duration
      return true;
    });
    // gantt.showLightbox = function(id) {
    //   var task = gantt.getTask(id);
    //   console.log('task', task);
    //   return true;
    //   };
  }

  componentDidMount() {
    gantt.init(this.ganttContainer);
    this.initGanttEvents();
    gantt.clearAll();
    gantt.parse(this.props.tasks);
    this.initGanttChart()
    currentScope = this;
  }

  initGanttChart() {
    gantt.config.columns = [
      { name: "add", label: "", width: 30 },
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
    // gantt.config.readonly = true;
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

    gantt.config.lightbox.sections = [
      { name: "title", height: 30, map_to: "text", type: "textarea", focus: true },
      { name: "description", height: 38, map_to: "body", type: "textarea" },
      {
        name: "status", height: 25, map_to: "status", type: "select", options: this.props.statusList
      },
      {
        name: "priority", height: 25, map_to: "priority", type: "select", options: [
          { key: "High", label: "High" },
          { key: "Normal", label: "Normal" },
          { key: "Low", label: "Low" }
        ]
      },
      {
        name: "type", height: 25, map_to: "type", type: "select", options: [
          { key: "Task", label: "Task" },
          { key: "Task2", label: "Milestone" }
        ]
      },
      {
        name: "owner", height: 25, map_to: "owner", type: "select", options: this.props.teamMembers
      },
      { name: "time", height: 50, map_to: "auto", type: "duration" },


    ];
    gantt.locale.labels.section_title = "Title";
    gantt.locale.labels.section_status = "Status";
    gantt.locale.labels.section_priority = "Priority";
    gantt.locale.labels.section_type = "Type";
    gantt.locale.labels.section_owner = "Assigned To";
  }

  render() {
    this.setZoom(this.props.zoom);
    return (
      <div
        ref={(input) => { this.ganttContainer = input }}
        style={{ width: '100%', height: '100%' }}
      ></div>
    );
  }
}