import { Project, Tag, File } from './Project';
export interface IProjectLevelDashboardState {
    description?: string;
    project: Project;
    tagList : Array<Tag>;
    attachment: File;
  }