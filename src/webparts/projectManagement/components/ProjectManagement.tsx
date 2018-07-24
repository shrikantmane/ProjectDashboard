import * as React from 'react';
import "primereact/resources/primereact.min.css";
import "primeicons/primeicons.css";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from './ProjectManagement.module.scss';
import { IProjectManagementProps } from './IProjectManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import ProjectListTable from './ProjectList/ProjectListTable';

export default class ProjectManagement extends React.Component<IProjectManagementProps, {}> {
  public render(): React.ReactElement<IProjectManagementProps> {
    // return (
    //   <div className={ styles.projectManagement }>
    //     <div className={ styles.container }>
    //       <div className={ styles.row }>
    //         <div className={ styles.column }>
    //           <span className={ styles.title }>Welcome to SharePoint!</span>
    //           <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
    //           <p className={ styles.description }>{escape(this.props.description)}</p>
    //           <a href="https://aka.ms/spfx" className={ styles.button }>
    //             <span className={ styles.label }>Learn more</span>
    //           </a>
    //         </div>
    //       </div>
    //     </div>
    //   </div>
    // );
    return (
      <div>
        <ProjectListTable></ProjectListTable>
      </div>
    );
  }
}
