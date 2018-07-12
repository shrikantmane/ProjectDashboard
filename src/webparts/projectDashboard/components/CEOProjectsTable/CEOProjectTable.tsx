import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";

import { ICEOProjectProps } from './ICEOProjectProps';
import { ICEOProjectState } from './ICEOProjectState';
import { CEOProjects } from "./CEOProject";
import CEOProjectTimeLine from '../CEOProjectTimeLine/CEOProjectTimeLine';
import ProjectTimeLine  from "../CEOProjectTimeLine/ProjectTimeLine";

export default class CEOProjectTable extends React.Component<ICEOProjectProps, ICEOProjectState> {
  constructor(props) {
    super(props);
    this.state = {
      projectList : new Array<CEOProjects>(),
      projectTimeLine : new Array<ProjectTimeLine>()
    }
  }

  componentDidMount(){
     this.getProjectList();
  }
  componentWillReceiveProps(nextProps){
  }

/* Private Methods */

  statusTemplate(rowData : CEOProjects, column){
    if(rowData.Status0)
     return (<div style={{backgroundColor: rowData.Status0.Status_x0020_Color, height: '2.9em', width:'100%', textAlign: 'center', paddingTop: 7, color: '#fff'}}>{rowData.Status0.Status}</div>)
  }

  /* Html UI */ 
  public render(): React.ReactElement<ICEOProjectProps> {
    var tasks = [
      {
            start: '2018-10-01',
            end: '2018-10-08',
            name: 'Redesign website',
            id: "Task 0",
            progress: 20
        },
        {
            start: '2018-10-03',
            end: '2018-10-06',
            name: 'Write new content',
            id: "Task 1",
            progress: 5,
            dependencies: 'Task 0'
        },
        {
            start: '2018-10-04',
            end: '2018-10-08',
            name: 'Apply new styles',
            id: "Task 2",
            progress: 10,
            dependencies: 'Task 1'
        },
        {
            start: '2018-10-08',
            end: '2018-10-09',
            name: 'Review',
            id: "Task 3",
            progress: 5,
            dependencies: 'Task 2'
        },
        {
            start: '2018-10-08',
            end: '2018-10-10',
            name: 'Deploy',
            id: "Task 4",
            progress: 0,
            dependencies: 'Task 2'
        },
        {
            start: '2018-10-11',
            end: '2018-10-11',
            name: 'Go Live!',
            id: "Task 5",
            progress: 0,
            dependencies: 'Task 4',
            custom_class: 'bar-milestone'
        },
        {
          start: '2018-10-12',
          end: '2018-10-13',
          name: 'Go Live!',
          id: "Task 6",
          progress: 0,
          dependencies: 'Task 5',
          custom_class: 'bar-milestone'
      },
      {
        start: '2018-10-13',
        end: '2018-10-15',
        name: 'Go Live!',
        id: "Task 7",
        progress: 0,
        dependencies: 'Task 6',
        custom_class: 'bar-milestone'
    },
    {
      start: '2018-10-11',
      end: '2018-10-11',
      name: 'Go Live!',
      id: "Task 8",
      progress: 0,
      dependencies: 'Task 7',
      custom_class: 'bar-milestone'
  },
  {
    start: '2018-10-11',
    end: '2018-10-11',
    name: 'Go Live!',
    id: "Task 9",
    progress: 0,
    dependencies: 'Task 8',
    custom_class: 'bar-milestone'
  },
  {
    start: '2018-10-11',
    end: '2018-10-11',
    name: 'Go Live!',
    id: "Task 10",
    progress: 0,
    dependencies: 'Task 9',
    custom_class: 'bar-milestone'
  } 
  ];
    return (
      <div>
        <CEOProjectTimeLine tasks= {tasks} ></CEOProjectTimeLine>
        <div>
          <DataTable value={this.state.projectList} responsive={true}>          
              <Column field="Project_x0020_ID" header="ID" />  
              <Column field="Project" header="Name" />
              <Column field="Priority" header="Priority" />
              <Column
              field="Status"
              header="Status"
              body={this.statusTemplate}
            />            
          </DataTable>
        </div>
      </div>
    );
  }

  /* Api Call*/

  private getProjectList(): void {
    
    sp.web.lists.getByTitle("Project")
    .items
    .select("Project_x0020_ID","Project", "StartDate", "DueDate", "AssignedTo/Title", "AssignedTo/ID","Status0/ID","Status0/Status","Status0/Status_x0020_Color","Priority").expand("AssignedTo", "Status0")
    .getAll()
    .then((response) => {
      this.setState({projectList : response})
    }); 
    
  }
}
