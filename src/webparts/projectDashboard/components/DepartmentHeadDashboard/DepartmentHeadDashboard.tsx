import * as React from 'react';
import { IDepartmentHeadDashboardProps } from './IDepartmentHeadDashboardProps';
import { IDepartmentHeadDashboardState } from './IDepartmentHeadDashboardState';

export default class DepartmentHeadDashboard extends React.Component<IDepartmentHeadDashboardProps, IDepartmentHeadDashboardState> {
  
    public render(): React.ReactElement<IDepartmentHeadDashboardProps> {
    return (
      <div>
         <h6>Department Head Dashboard</h6>
      </div>
    );
  }
}
