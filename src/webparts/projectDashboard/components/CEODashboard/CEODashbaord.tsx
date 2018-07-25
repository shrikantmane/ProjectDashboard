import * as React from 'react';
import { ICEODashboardProps } from './ICEODashboardProps';
import { ICEODashboardState } from './ICEODashboardState';
import CEOProjectTable from '../CEOProjectsTable/CEOProjectTable';
import CEOProjectTimeLine from '../CEOProjectTimeLine/CEOProjectTimeLine';

export default class CEODashboard extends React.Component<ICEODashboardProps, ICEODashboardState> {
  
    public render(): React.ReactElement<ICEODashboardProps> {
    return (
      <div>
        <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable>
      </div>
    );
  }
}
