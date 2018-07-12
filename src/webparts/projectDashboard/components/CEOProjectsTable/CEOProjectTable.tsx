import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";

import { ICEOProjectProps } from './ICEOProjectProps';
import { ICEOProjectState } from './ICEOProjectState';
import { CEOProjects } from "./CEOProject";

export default class CEOProjectTable extends React.Component<ICEOProjectProps, ICEOProjectState> {
  constructor(props) {
    super(props);
    this.state = {
      projectList : [],
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
    return (
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
