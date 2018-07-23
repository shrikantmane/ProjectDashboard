import { CEOProjects } from "./CEOProject";
import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export interface ICEOProjectState {
  projectList: Array<CEOProjects>,
  projectTimeLine: Array<ProjectTimeLine>,
  expandedRows? :any,
  globalFilter?: any,
  projectName: string,
  ownerName: string,
  status: string,
  priority: string,
  isLoading: boolean,
  isTeamMemberLoaded : boolean,
  isKeyDocumentLoaded :boolean,
  isTagLoaded :boolean,
  expandedRowID : number
}