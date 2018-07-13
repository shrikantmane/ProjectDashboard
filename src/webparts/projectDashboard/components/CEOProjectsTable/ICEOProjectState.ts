import { CEOProjects } from "./CEOProject";
import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export interface ICEOProjectState {
  projectList: Array<CEOProjects>,
  projectTimeLine: Array<ProjectTimeLine>,
  expandedRows? :any
}