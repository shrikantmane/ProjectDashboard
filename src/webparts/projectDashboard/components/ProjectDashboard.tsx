import * as React from 'react';
import styles from './ProjectDashboard.module.scss';
import 'primereact/resources/primereact.min.css';
import 'primeicons/primeicons.css';

import { IProjectDashboardProps } from './IProjectDashboardProps';
import CEODashboard from './CEODashboard/CEODashbaord';
export default class ProjectDashboard extends React.Component<IProjectDashboardProps, {}> {
  public render(): React.ReactElement<IProjectDashboardProps> {
    return (
      // <div className={ styles.projectDashboard }>
      //   <div className={ styles.container }>
      //     <div className={ styles.row }>
      //       <div className={ styles.column }>
      //         <span className={ styles.title }>Welcome to SharePoint!</span>
      //         <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
      //         <p className={ styles.description }>{escape(this.props.description)}</p>
      //         <a href="https://aka.ms/spfx" className={ styles.button }>
      //           <span className={ styles.label }>Learn more</span>
      //         </a>
      //       </div>
      //     </div>
      //   </div>
      // </div>
      <div>
        <CEODashboard></CEODashboard>
      </div>
    );
  }
}
