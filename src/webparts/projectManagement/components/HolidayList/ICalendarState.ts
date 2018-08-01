import { CalendarList} from "./CalendarList";
import { Calendar } from "office-ui-fabric-react/lib/Calendar";
//import { TeamMembers } from "../../../../../lib/webparts/projectDashboard/components/CEOProjectsTable/CEOProject";
// import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export interface ICalendarState {
    projectList: Array<CalendarList>,
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
    showComponent: boolean,
    informationID ?: number
}