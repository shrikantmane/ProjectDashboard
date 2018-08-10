import { Task } from "../DepartmentDashboard";
export interface IProjectTaskListState {
    description ?: string;
    taskList :  Array<Task>;
    showComponent: boolean,
    showDocumentComponent: boolean,
    taskID ?: number,
    documentID ?: number,
  }