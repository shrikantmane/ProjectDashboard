import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectViewProps } from "./IProjectViewProps";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "../ProjectManagement.module.scss";
import TeamListTable from '.././TeamMembers/TeamListTable';
import DocumentListTable from '.././ProjectDocuments/DocumentListTable'
import RequirementListTable from '.././ProjectRequirement/RequirementListTable'
import InformationListTable from '.././ProjectInformation/InformationListTable'
import CalendarListTable from '.././HolidayList/CalendarListTable'
import ProjectPlan from '.././ProjectPlan/ProjectPlan'
//import FullCalendar from 'fullcalendar-reactwrapper';
import { IProjectViewState } from './IProjectViewState';
import CalendarViewListTable from '.././HolidayCalendar/CalendarViewListTable'
import AddEvent from '../AddEvent/AddEvent';
import { SPComponentLoader } from "@microsoft/sp-loader";
import FullCalendar from 'fullcalendar-reactwrapper';
export default class ProjectViewDetails extends React.Component<
    IProjectViewProps,
    IProjectViewState
    > {
    constructor(props) {
        super(props);
        this.state = {
            project: "",
            startdate: "",
            enddate: "",
            onhold: "",
            owner: [],
            priority: "",
            complexity: "",
            status: "",
            statuscolor: "",
            informationlist: "",
            documentlist: "",
            requirementlist: "",
            teammemberlist: "",
            Id: "",
            calendarList: "",
            showComponent: false,
            refreshCalender: false,
            imgURL: "",
            Risks: "",
            events: [
                {
                    id: 10,
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
            scheduleList : ""
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.onRefreshCalender = this.onRefreshCalender.bind(this)
    }

    violet: any;
    red: any;
    componentDidMount() {
        console.log('componentDidMount');
        const { match: { params } } = this.props;
        SPComponentLoader.loadCss(
            "https://cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.9.0/fullcalendar.css"
        );
        let id = params.id.split('_')[0];
        let scrollTo = params.id.split('_')[1] ? params.id.split('_')[1] : '';
        this.getProjectData(id);
        var elmnt = document.getElementById(scrollTo);
        if (elmnt)
            elmnt.scrollIntoView();
    }
    onRefreshCalender(refresh) {
        this.setState({ refreshCalender: true });
    }
    getProjectData(id) {
        console.log('params : ' + id);
        sp.web.lists.getByTitle("Project").items
            .select("ID", "Project", "DueDate", "Priority", "On_x0020_Hold_x0020_Status", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "AssignedTo/EMail", "Priority", "Task_x0020_List", "Project_x0020_Team_x0020_Members", "Project_x0020_Document", "Requirements", "Project_x0020_Infromation", "Project_x0020_Calender", "StartDate", "Risks","Schedule_x0020_List").expand("Status0", "AssignedTo")
            .filter('ID eq \'' + id + '\'')
            .getAll()
            .then((response) => {
                let status;
                if (response != null) {

                    console.log('Project by names', response);
                     if (response[0].On_x0020_Hold_x0020_Status === true)
                        status = "Yes";
                    else
                        status = "No";
                    this.setState({
                        project: response ? response[0].Project : '',
                        startdate: response[0].StartDate ? new Date(response[0].StartDate).toDateString() : '',
                        enddate: response[0].DueDate ? new Date(response[0].DueDate).toDateString() : '',
                        onhold: response ? status : null,
                        owner: response ? (response[0].AssignedTo!==undefined ?(response[0].AssignedTo? response[0].AssignedTo[0].Title: null) :null) : null,
                        priority: response ? response[0].Priority : '',
                        status: response ? (response[0].Status0===undefined ?(response[0].Status0? response[0].Status0.Status: null) : response[0].Status0.Status ) : null,
                        //status: response ? (response[0].Status0 ? response[0].Status0.Status : null) : '',
                        informationlist: response ? response[0].Project_x0020_Infromation : '',
                        teammemberlist: response ? response[0].Project_x0020_Team_x0020_Members : '',
                        requirementlist: response ? response[0].Requirements : '',
                        documentlist: response ? response[0].Project_x0020_Document : '',
                        Id: response ? response[0].ID : '',
                        calendarList: response ? response[0].Project_x0020_Calender : '',
                        statuscolor: response ? (response[0].Status0!==undefined ?(response[0].Status0? response[0].Status0.Status_x0020_Color: null) : null) : null,
                        //statuscolor: response ? (response[0].Status0 ? response[0].Status0.Status_x0020_Color : null) : null,
                        imgURL:response ? (response[0].AssignedTo!==undefined ?(response[0].AssignedTo? "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" + response[0].AssignedTo[0].EMail + "&UA=0&size=HR64x64&sc=1531997060853" : null) : null) : null,
                        Risks: response ? response[0].Risks : '',
                        scheduleList : response[0] ? response[0].Schedule_x0020_List : ""

                    });
                    console.log("helllo", this.state)
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
        console.log('render');
        return (
            <div className="PanelContainer">
                <div className="viewProject">
                    <section className="main-content-section">
                        <div className="wrapper">
                            <div className="row">
                                <div className="col-12 col-sm-12">
                                    <div className="well headerWell">
                                        <div className="row">
                                            <div className="col-sm-2 col-12">
                                                <div className="current-project-conatiner">
                                                    <p id="project-name">{this.state.project}</p>
                                                    <img className="profile-img-style" src={this.state.imgURL} style={{ display: !this.state.imgURL? 'none': '' }}  />
                                                    <p>{this.state.owner}</p>
                                                </div>
                                            </div>
                                            <div className="col-12 col-sm-10">
                                                <div className="row">
                                                    <div className="col-12 col-sm-2">
                                                        <div className="current-project-conatiner">
                                                            <span>Start Date</span>
                                                            <span>{this.state.startdate}</span>
                                                        </div>
                                                    </div>
                                                    <div className="col-12 col-sm-2">
                                                        <div className="current-project-conatiner">
                                                            <span>End Date</span>
                                                            <span>{this.state.enddate}</span>
                                                        </div>
                                                    </div>
                                                    <div className="col-12 col-sm-2">
                                                        <div className="current-project-conatiner">
                                                            <span>On Hold</span>
                                                            <div>{this.state.onhold}</div>  
                                                        </div>
                                                    </div>
                                                    <div className="col-12 col-sm-2">
                                                        <div className="current-project-conatiner">
                                                            <span>Priority</span>
                                                            <span className="tags grey">{this.state.priority}</span>
                                                        </div>
                                                    </div>

                                                    <div className="col-12 col-sm-2">
                                                        <div className="current-project-conatiner">
                                                            <span>Complexity</span>
                                                            <span className="tags grey">{this.state.Risks}</span>
                                                        </div>
                                                    </div>
                                                    <div className="col-12 col-sm-2">
                                                        <div className="current-project-conatiner">
                                                            <span>Status</span>
                                                            <span className="tags orange" style={{ color: this.state.statuscolor, border: "1px solid " + this.state.statuscolor }}>{this.state.status}</span>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                <div className="col-12 col-sm-6 half-list-left" id="member">
                                    <div className="well hbc-list hbc-list-member">
                                        <div><TeamListTable list={this.state.teammemberlist} projectId={this.state.Id}></TeamListTable></div>
                                    </div>
                                </div>
                                <div className="col-12 col-sm-6 half-list-right" id="document">
                                    <div className="well hbc-list hbc-list-document">
                                        <div><DocumentListTable list={this.state.documentlist} projectId={this.state.Id}></DocumentListTable></div>
                                    </div>
                                </div>
                                
                                <div className="col-12 col-sm-12">
                                    <div className="well hbc-list hbc-list-roles">
                                        <div><InformationListTable list={this.state.informationlist} projectId={this.state.Id}></InformationListTable></div>
                                    </div>
                                </div>

                                <div className="col-12 col-sm-12">
                                    <div className="well">
                                        <div><ProjectPlan scheduleList={this.state.scheduleList}></ProjectPlan></div>
                                    </div>
                                </div>

                                <div className="col-12 col-sm-12" id="requirement">
                                    <div className="well hbc-list hbc-list-requirement">
                                        <div><RequirementListTable list={this.state.requirementlist} projectId={this.state.Id}></RequirementListTable></div>

                                    </div>
                                </div>
                                <div className="col-12 col-sm-6 half-list-left" ref={(section) => { this.violet = section; }}>
                                    <div className="well  hbc-list-holidaylist">
                                        <div><CalendarListTable list={this.state.calendarList} projectId={this.state.Id} onRefreshCalender={this.onRefreshCalender}></CalendarListTable></div>
                                    </div>
                                </div>
                                <div className="col-12 col-sm-6 half-list-right" id="content2">
                                    <div className="well hbc-list-holidaycalender">
                                        <div className="content-section implementation">
                                            <h5>Holiday Calender</h5>
                                        </div>
                                        <div className="col-sm-12 col-12 cardHeading">
                                            <div className="content-section implementation">
                                                {/* <h5>Events</h5> */}
                                                {/* <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px",borderRadius: "30px",backgroundColor: "#0078d7", border: "1px solid #0078d7",padding: "4px 10px 6px",color: "white",marginLeft: "20px"}} onClick={this.onAddProject}> Add Holiday </button>
                                    {this.state.showComponent ?
                                    <AddEvent list={this.state.calendarList} projectId={this.state.Id} /> :
                                    null
                                    } */}
                                                <div  >
                                                    {/* <FullCalendar

                                                          header={{
                                                          left: 'prev,next today myCustomButton',
                                                          center: 'title',
                                                          right: 'month,agendaWeek,agendaDay,listWeek'
                                                          }}
                                                          navLinks={true} // can click day/week names to navigate views
                                                          editable={true}
                                                          eventLimit={true} // allow "more" link when too many events
                                                          events={this.state}
                                                          /> */}
                                                    <CalendarViewListTable list={this.state.calendarList} projectId={this.state.Id} refreshCalender={this.state.refreshCalender}></CalendarViewListTable>
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

            // <div>
            //     <button onClick={this.scrollToTopWithCallback}>Go To Violet</button>
            //      <section id="content1" style={{height : "1000px", backgroundColor: "yellow", width:"500px"}} ref={(section) => { this.violet = section; }}>yellow</section>
            //     <section id="content2" style={{height : "1000px", backgroundColor: "red", width:"500px"}}  ref={(section) => { this.red = section; }}>Red</section>

            //     </div>
        );
    }
}
