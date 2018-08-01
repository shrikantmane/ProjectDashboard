import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectViewProps } from "./IProjectViewProps";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "../ProjectManagement.module.scss";
import TeamListTable from '.././TeamMembers/TeamListTable';
import DocumentListTable from '.././ProjectDocuments/DocumentListTable'
import RequirementListTable from '.././ProjectRequirement/RequirementListTable'
import InformationListTable from '.././ProjectInformation/InformationListTable'
import FullCalendar from 'fullcalendar-reactwrapper';
import {IProjectViewState} from './IProjectViewState';
import AddEvent from '../AddEvent/AddEvent';
export default class ProjectViewDetails extends React.Component<
    IProjectViewProps,
    IProjectViewState
    > {
    constructor(props) {
        super(props);
        this.state = {
            informationlist:"",
            documentlist:"",
            requirementlist:"",
            teammemberlist:"",
            Id:"",
            calendarList:"",
            showComponent:false,
            events:[
                            {id:10,
                              title: 'All Day Event',
                                  start: '2018-07-01'
                              }
                            //  {
                            //      title: 'Long Event',
                            //      start: '2018-08-07',
                            //    end: '2018-08-10'
                            //  },
                            //   {
                            //       id: 999,
                            //       title: 'Repeating Event',
                            //       start: '2018-07-09T16:00:00'
                            //   },
                            //    {
                            //   id: 999,
                            //       title: 'Repeating Event',
                            //       start: '2018-07-16T16:00:00'
                            //   },
                            //   {
                            //       title: 'Conference',
                            //       start: '2018-08-11',
                            //       end: '2018-08-13'
                            //   },
                            //    {
                            //       title: 'Meeting',
                            //       start: '2018-07-12T10:30:00',
                            //       end: '2018-07-12T12:30:00'
                            //   },
                            //   {
                            //       title: 'Birthday Party',
                            //       start: '2018-07-27T07:00:00'
                            //    },
                            //    {
                            //       title: 'Click for Google',
                            //       start: '2018-07-30'
                            //   }
                          ],		
        };
        this.onAddProject=this.onAddProject.bind(this);
    }
    componentDidMount() {
        const {
      match: { params }
    } = this.props;
        console.log('params : ' + params.id);
        sp.web.lists.getByTitle("Project").items
        .select("ID","Project", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "Priority","Task_x0020_List","Project_x0020_Team_x0020_Members","Project_x0020_Document","Requirements","Project_x0020_Infromation","Project_x0020_Calender").expand("Status0", "AssignedTo")
        .filter('ID eq \'' + params.id + '\'')
        .getAll()
        .then((response) => {
        if(response!=null){
          console.log('Project by names', response);
         
          this.setState({
            informationlist : response[0].Project_x0020_Infromation,
            documentlist : response[0].Project_x0020_Document,
            requirementlist : response[0].Requirements,
            teammemberlist:response[0].Project_x0020_Team_x0020_Members,
            Id: response[0].ID,
            calendarList:response[0].Project_x0020_Calender
          });
          console.log("helllo",this.state)
         }

         
      }).catch((e: Error) => {
          alert(`There was an error : ${e.message}`);
        });
    }
    onAddProject() {
        console.log('button clicked');
         this.setState({
         showComponent: true,
         });
    }
    

    public render(): React.ReactElement<IProjectViewProps> {
        return (
            <section className="main-content-section">
                <div className="wrapper">
                    <div className="row">
                        <div className="col-lg-12 col-md-12 col-sm-12">
                            <section id="step1">
                                <div className="well">
                                    <div className="row">
                                        <div className="clearfix"></div>
                                        <div className="row portfolioNTasks">
                                            {/* first box */}
                                            <div className="col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <div className="well recommendedProjects">
                                                    <div className="row">
                                                        <div className="col-sm-12 col-md-12 col-lg-12 cardHeading">
                                                            <div> 
                                                                <h5 className="pull-left heading-style">Team Members</h5>
                                                                <div><TeamListTable list={this.state.teammemberlist} projectId={this.state.Id}></TeamListTable></div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <div className="well recommendedProjects">
                                                    <div className="row">
                                                        <div className="col-sm-12 col-md-12 col-lg-12 cardHeading">
                                                            <div> 
                                                                <h5 className="pull-left heading-style">Documents</h5>
                                                                <div><DocumentListTable list={this.state.documentlist} projectId={this.state.Id}></DocumentListTable></div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                                                <div className="well recommendedProjects">
                                                    <div className="row">
                                                        <div className="col-sm-12 col-md-12 col-lg-12 cardHeading">
                                                            <div> 
                                                                <h5 className="pull-left heading-style">Requirements</h5>
                                                                <div><RequirementListTable list={this.state.requirementlist} projectId={this.state.Id}></RequirementListTable></div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <div className="well recommendedProjects">
                                                    <div className="row">
                                                        <div className="col-sm-12 col-md-12 col-lg-12 cardHeading">
                                                            <div> 
                                                                <h5 className="pull-left heading-style">Responsibilities</h5>
                                                                <div><InformationListTable list={this.state.informationlist} projectId={this.state.Id}></InformationListTable></div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <div className="well recommendedProjects">
                                                    <div className="row">
                                                        <div className="col-sm-12 col-md-12 col-lg-12 cardHeading">
                                                            <div> 
                                                                <h5 className="pull-left heading-style">Events</h5>
                                                                <div > <button onClick={this.onAddProject}>Add Event </button></div>
                                                                {this.state.showComponent ?
                                                           <AddEvent list={this.state.calendarList} projectId={this.state.Id}  /> :
                                                                 null
                                                                                         }
                                                                <div >
                                                            <FullCalendar
            
                                                               header = {{
                                                                        left: 'prev,next today myCustomButton',
                                                                    center: 'title',
                                                                        right: 'month,agendaWeek,agendaDay,listWeek'
                                                                    }}
                                                                    navLinks= {true} // can click day/week names to navigate views
                                                                    editable= {true}
                                                                    eventLimit= {true} // allow "more" link when too many events
                                                                events={this.state}
                                                                />

                                                              </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </section>
                        </div>
                    </div>
                </div>
            </section>
        );
    }
}
