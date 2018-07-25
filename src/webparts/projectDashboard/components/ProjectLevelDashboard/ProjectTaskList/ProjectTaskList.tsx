import * as React from 'react';
import { IProjectTaskListProps } from './IProjectTaskListProps';
import { IProjectTaskListState } from './IProjectTaskListState';

export default class ProjectTaskList extends React.Component<IProjectTaskListProps, IProjectTaskListState> {
  
    public render(): React.ReactElement<IProjectTaskListProps> {
    return (
      <div>
        {/* <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable> */}
      </div>
    );
  }
}
