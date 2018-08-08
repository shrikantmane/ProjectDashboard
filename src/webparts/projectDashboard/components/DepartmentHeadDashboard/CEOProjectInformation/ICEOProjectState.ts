import { Projects } from "../DepartmentDashboard";;
import { ProjectTimeLine }  from "../TimeLineChart/TimeLine";

export interface ICEOProjectState {
  projectList: Array<Projects>,
  projectTimeLine: ProjectTimeLine,
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
  expandedRowID : number,
}