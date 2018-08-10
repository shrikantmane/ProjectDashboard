import * as React from 'react';
import { DataTable } from "primereact/components/datatable/DataTable";
import { ICEODashboardProps } from './ICEODashboardProps';
import { ICEODashboardState } from './ICEODashboardState';
import CEOProjectInformation from './CEOProjectInformation/CEOProjectInformation';
import CEOProjectTimeLine from './CEOProjectTimeLine/CEOProjectTimeLine';

export default class CEODashboard extends React.Component<ICEODashboardProps, ICEODashboardState> {
  
    public render(): React.ReactElement<ICEODashboardProps> {
    return (
      <div>
        <CEOProjectInformation webPartTitle={this.props.webPartTitle}></CEOProjectInformation>
      </div>
    );
  }
}
