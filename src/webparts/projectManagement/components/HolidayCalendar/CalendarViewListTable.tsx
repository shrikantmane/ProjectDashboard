import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { ICalendarViewProps } from "./ICalendarViewProps";
import { ICalendarViewState } from "./ICalendarViewState";
import {
    CalendarListView
} from "./CalendarViewList";
import FullCalendar from 'fullcalendar-reactwrapper';
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';

export default class ProjectListTable extends React.Component<
    ICalendarViewProps,
    ICalendarViewState
    > {
    constructor(props) {
        super(props);
        // this.state = {
        //    projectList: new Array<Project>(),
        //     //   projectTimeLine: new Array<ProjectTimeLine>(),
        //     projectName: null,
        //     ownerName: null,
        //     status: null,
        //     priority: null,
        //     isLoading: true,
        //     isTeamMemberLoaded: false,
        //     isKeyDocumentLoaded: false,
        //     isTagLoaded: false,
        //     expandedRowID: -1,
        //     expandedRows: []
        // };
        this.state = {
            projectList: new Array<CalendarListView>(),
            showComponent: false,
            selectedFile: "",
            documentID: "",
            events: [
                // {
                //     id: 10,
                //     title: 'All Day Event',
                //     start: '2018-07-01'
                // }
            ],
        }
        this.refreshGrid = this.refreshGrid.bind(this);

    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if (this.props.list != "" || this.props.list != null) {
            this.getProjectCalendar(this.props.list);
        }

    }
    refreshGrid() {
        this.getProjectCalendar(this.props.list);
    }


    private deleteListItem(rowData, e): any {
        e.preventDefault();
        console.log('Edit :' + rowData);
        sp.web.lists.getByTitle(this.props.list).
            items.getById(rowData.ID).delete().then((response) => {
                console.log(this.props.list + ` item deleted`);
                this.getProjectCalendar(this.props.list);
            });
    }
    componentWillReceiveProps(nextProps) {
        if (nextProps.list != "" || nextProps.list != null) {
            this.getProjectCalendar(nextProps.list);
        }
    }

    /* Private Methods */

    /* Html UI */



    public render(): React.ReactElement<ICalendarViewState> {
        return (
           
                <div className="col-xs-12 col-sm-12 col-md-12 col-lg-12" >

                    <FullCalendar

                        header={{
                            left: 'prev,next today myCustomButton',
                            center: 'title',
                            right: 'month,agendaWeek,agendaDay,listWeek'
                        }}
                        navLinks={true} // can click day/week names to navigate views
                        // editable={true}
                        eventLimit={true} // allow "more" link when too many events
                        events={this.state.events}
                    />
                </div>
        );
    }

    /* Api Call*/



    getProjectCalendar(list) {
        if ((list) != "") {
            sp.web.lists.getByTitle(list).items
                .select("ID", "Title", "EndDate", "EventDate")
                .get()
                .then((response) => {
                    let tempArray = {};
                    let tempList = [];
                    response.forEach(element => {
                        var date = new Date(element.EndDate);
                  var date1:any

                  date1=  date.setDate(date.getDate() + 1);
                        tempArray = {
                            title: element.Title, start:element.EventDate.split('T')[0], end: date1
                        }
                        tempList.push(tempArray);
                        console.log('information by name', response);
                        this.setState({ events: tempList });
                       
                    });
                })
        }
    }

    getProjectDocuments(list) {
        if ((list) != "") {

            sp.web.lists.getByTitle(list).items
                .select("ID", "File", "Author/ID", "Author/Title", "Created").expand("File", "Author")

                .get()
                .then((response) => {
                    console.log('calendar by name', response);
                    this.setState({ projectList: response });
                });
        }
    }




}