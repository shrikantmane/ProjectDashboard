/*global gantt*/
import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { find} from "lodash";
// import 'dhtmlx-gantt';
import 'dhtmlx-gantt';
import 'dhtmlx-gantt/codebase/dhtmlxgantt.css';
import 'dhtmlx-gantt/codebase/ext/dhtmlxgantt_tooltip.js';
import { Chart, ChartData, ChartLink } from "./Project";
declare var gantt: any;
var currentScope:any;

export default class Gantt extends React.Component<any, any>{
  ganttContainer: any;

  constructor(props) {
    super(props);
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

  initGanttEvents() {
    if(gantt.ganttEventsInitialized){
      return;
    }
    gantt.ganttEventsInitialized = true;

    gantt.attachEvent('onAfterTaskAdd', (id, task: any) => {
      // if(this.props.onTaskUpdated) {
      //   this.props.onTaskUpdated(id, 'inserted', task);
      // }
      
      console.log('task22222222222', task);
      let duration = task.type === "Task2" ? 0 : task.duration;
      let status: any = find(this.props.statusList, { key : task.status });
      let owner : any  = find(this.props.teamMembers, { key : task.owner });
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
        console.log("ScheduleTask Added !");
      }).catch(err => {
        console.log("Error while adding ScheduleTask", err);
      });
    });

    gantt.attachEvent('onAfterTaskUpdate', (id, task) => {
      // if(this.props.onTaskUpdated) {
      //   this.props.onTaskUpdated(id, 'updated', task);
      // }
      console.log('task1111111111', task);
      let duration = task.type === "Task2" ? 0 : task.duration;
      let status: any = find(this.props.statusList, { key : task.status });
      let owner : any  = find(this.props.teamMembers, { key : task.owner });
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
        console.log("ScheduleTask Updated !");
      }).catch(err => {
        console.log("Error while updating ScheduleTask", err);
      });
    });

    gantt.attachEvent('onAfterTaskDelete', (id, task) => {
      // if(this.props.onTaskUpdated) {
      //   this.props.onTaskUpdated(id, 'deleted');
      // }
      console.log('id4364634634', id , task);
      sp.web.lists.getByTitle(this.props.scheduleList).items.getById(task.id).delete().then(_ => {
        console.log("ScheduleTask deleted !");
      }).catch(err => {
        console.log("Error while deleting ScheduleTask", err);
      });
  
    });

    gantt.attachEvent('onAfterLinkAdd', (id, link) => {
      if(this.props.onLinkUpdated) {
        this.props.onLinkUpdated(id, 'inserted', link);
      }
    });

    gantt.attachEvent('onAfterLinkUpdate', (id, link) => {
      if(this.props.onLinkUpdated) {
        this.props.onLinkUpdated(id, 'updated', link);
      }
    });

    gantt.attachEvent('onAfterLinkDelete', (id, link) => {
      if(this.props.onLinkUpdated) {
        this.props.onLinkUpdated(id, 'deleted');
      }
    });
  //   gantt.attachEvent("onTaskCreated", function(id, task){
  //     console.log('onTaskCreated',task);
  //     task.statusBackgroudColor = "blue";
  //     return true;
  // });
  gantt.attachEvent("onLightboxSave", function(id, item){
    console.log('onLightboxSaveitem', item)
    let status: any = find(currentScope.props.statusList, { key : item.status });
    item.statusBackgroudColor = status.color;
    item.actualDuration =  item.type === "Task2" ? 0 : item.duration;
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
    console.log('this.props.teamMembers', this.props.teamMembers);
    gantt.init(this.ganttContainer);
    this.initGanttEvents();
    gantt.clearAll();
    gantt.parse(this.props.tasks);
    this.initGanttChart()
    currentScope = this;
  }

  initGanttChart() {
    gantt.config.columns = [
      {name:"add", label:"", width:30 },
      {
        name: "attachment", label: "", width: 30,
        template: function (obj) {
          return "<a href=''><i class='fas fa-paperclip'></i></a>";
        }
      },
      {
        name: "comments", label: "", width: 30,
        template: function (obj) {
          return "<a href=''><i class='far fa-comments'></i></a>";
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
          console.log('actualDuration', obj.actualDuration);
          return ("<span>" + obj.actualDuration + "</span>")
        }
      },
    ];
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
      console.log('task_class', task.actualDuration);
      if (task.actualDuration == 0) {
        return "milestone";
      }
    };

    gantt.config.lightbox.sections = [
      {name:"title", height:30, map_to:"text", type:"textarea",focus:true},
      {name:"description", height:38, map_to:"body", type:"textarea"},
      {
        name: "status", height: 25, map_to: "status", type: "select", options: this.props.statusList
      },
      {
        name: "priority", height: 25, map_to: "priority", type: "select", options: [
          { key: "High", label: "High" },
          { key:  "Normal" , label: "Normal" },
          { key:"Low", label: "Low" }
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
      {  name:"time", height:50, map_to:"auto", type:"duration"},

      
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