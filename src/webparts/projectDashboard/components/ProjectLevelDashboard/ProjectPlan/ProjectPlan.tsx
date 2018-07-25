import * as React from 'react';
import { IProjectPlanProps } from './IProjectPlanProps';
import { IProjectPlanState } from './IProjectPlanState';

export default class ProjectPlan extends React.Component<IProjectPlanProps, IProjectPlanState> {
  
    public render(): React.ReactElement<IProjectPlanProps> {
    return (
      <div>
        {/* <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable> */}
      </div>
    );
  }
}
