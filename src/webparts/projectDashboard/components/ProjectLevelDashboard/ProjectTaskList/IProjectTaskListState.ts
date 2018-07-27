import { Task } from '../Project';
export interface IProjectTaskListState {
    description ?: string;
    taskList :  Array<Task>;
  }