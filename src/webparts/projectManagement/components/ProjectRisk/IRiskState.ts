import { Risk } from "./RiskList";
//import { TeamMembers } from "../../../../../lib/webparts/projectDashboard/components/CEOProjectsTable/CEOProject";
// import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export interface IRiskState {
    projectList: Array<Risk>,
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
}