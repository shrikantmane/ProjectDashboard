import { Task } from "../DepartmentDashboard";
export interface IProjectTaskListState {
    description ?: string;
    taskList :  Array<Task>;
    sortField: string;
    sortOrder: number;
  }