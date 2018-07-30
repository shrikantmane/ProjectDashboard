import { TeamMembers } from "./TeamList";
//import { TeamMembers } from "../../../../../lib/webparts/projectDashboard/components/CEOProjectsTable/CEOProject";
// import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export interface ITeamState {
    projectList: Array<TeamMembers>,
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