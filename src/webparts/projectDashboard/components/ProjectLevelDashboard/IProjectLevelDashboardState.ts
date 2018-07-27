import { Project, Tag } from './Project';
export interface IProjectLevelDashboardState {
    description?: string;
    project: Project;
    tagList : Array<Tag>
  }