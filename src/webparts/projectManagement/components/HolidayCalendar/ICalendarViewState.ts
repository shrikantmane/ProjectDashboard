import { CalendarListView } from "./CalendarViewList";
//import { TeamMembers } from "../../../../../lib/webparts/projectDashboard/components/CEOProjectsTable/CEOProject";
// import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export interface ICalendarViewState {
    projectList: Array<CalendarListView>,
    //   projectTimeLine: Array<ProjectTimeLine>,
    //   expandedRows? :any,
    //   globalFilter?: any,
    //   projectName: string,
    //   ownerName: string,
    //   status: string,
    //   priority: string,
    //   isLoading: boolean,
    //   isTeamMemberLoaded : boolean,
    //   isKeyDocumentLoaded :boolean,
    //   isTagLoaded :boolean,
    //   expandedRowID : number,
    showComponent: boolean
    selectedFile:any
    documentID:any,
//     events:[
//         // {
//         // id?:number,
//         // title?:string,
//         // start?:string,
//         // end?:string,
//         // }
// ]
events:any;
    
}