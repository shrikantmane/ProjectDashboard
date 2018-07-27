import * as React from 'react';
import "primereact/resources/primereact.min.css";
import "primeicons/primeicons.css";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from './ProjectManagement.module.scss';
import { IProjectManagementProps } from './IProjectManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import ProjectListTable from './ProjectList/ProjectListTable';
import TeamListTable from './TeamMembers/TeamListTable';
import RiskListTable from './ProjectRisk/RiskListTable';
import InformationListTable from './ProjectInformation/InformationListTable'
import RequirementListTable from './ProjectRequirement/RequirementListTable'
import DocumentListTable from './ProjectDocuments/DocumentListTable'
import ScheduleListTable from './ProjectSchedule/ScheduleListTable'
import FullCalendar from 'fullcalendar-reactwrapper';
import { SPComponentLoader } from "@microsoft/sp-loader";
import ProjectViewDetails from './ViewProject/ProjectViewDetails';
import { Switch, Route } from 'react-router-dom';

export default class ProjectManagement extends React.Component<IProjectManagementProps, {}> {
  componentDidMount() {
    SPComponentLoader.loadCss(
        "https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.9.0/fullcalendar.css"
    );
    
 
    
}
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
    // this.state = {
    //   events:[
    //               {
    //                   title: 'All Day Event',
    //                   start: '2018-07-01'
    //               },
    //               {
    //                   title: 'Long Event',
    //                   start: '2018-08-07',
    //                   end: '2018-08-10'
    //               },
    //               {
    //                   id: 999,
    //                   title: 'Repeating Event',
    //                   start: '2018-07-09T16:00:00'
    //               },
    //               {
    //                   id: 999,
    //                   title: 'Repeating Event',
    //                   start: '2018-07-16T16:00:00'
    //               },
    //               {
    //                   title: 'Conference',
    //                   start: '2018-08-11',
    //                   end: '2018-08-13'
    //               },
    //               {
    //                   title: 'Meeting',
    //                   start: '2018-07-12T10:30:00',
    //                   end: '2018-07-12T12:30:00'
    //               },
    //               {
    //                   title: 'Birthday Party',
    //                   start: '2018-07-27T07:00:00'
    //               },
    //               {
    //                   title: 'Click for Google',
    //                   start: '2018-07-30'
    //               }
    //           ],		
    //   }
    // return (
    //   <div>
    //      <div >
    //     <FullCalendar
            
    //      header = {{
    //         left: 'prev,next today myCustomButton',
    //         center: 'title',
    //         right: 'month,agendaWeek,agendaDay,listWeek'
    //     }}
    //     navLinks= {true} // can click day/week names to navigate views
    //     editable= {true}
    //     eventLimit= {true} // allow "more" link when too many events
    //    events={this.state}
    // />

    //   </div>
      
    //     <DocumentListTable></DocumentListTable>
    //      <div>
    //     <InformationListTable></InformationListTable>
    //       </div> 
    //       <div>
    //     <RequirementListTable></RequirementListTable>
    //       </div> 
    //       <div>
    //     <RiskListTable></RiskListTable>
    //       </div>
    //       <div>
    //     <ScheduleListTable></ScheduleListTable>
    //       </div>
    //       <div>
    //     <TeamListTable></TeamListTable>
    //       </div>
    //       <div>
    //           <ProjectListTable></ProjectListTable>
    //           </div>
    //   </div>
    // )
    return (
      // <div>
      //   <ProjectListTable></ProjectListTable>
      // </div>
      <Switch>
        <Route exact path='/' component={ProjectListTable} />
        <Route path='/viewProjectDetails/:id' component={ProjectViewDetails} />
      </Switch>

    );
  }
}
