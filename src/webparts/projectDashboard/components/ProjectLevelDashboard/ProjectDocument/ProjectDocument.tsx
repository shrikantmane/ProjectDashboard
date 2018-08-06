import * as React from 'react';
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import { IProjectDocumentProps } from './IProjectDocumentProps';
import { IProjectDocumentState } from './IProjectDocumentState';
import { Document } from '../Project';
import moment from 'moment/src/moment';

export default class ProjectDocument extends React.Component<IProjectDocumentProps, IProjectDocumentState> {

  constructor(props) {
    super(props);
    this.state = {
      documentList: new Array<Document>()
    };
  }

  componentWillReceiveProps(nextProps) {
    if (this.props.projectDocument != nextProps.projectDocument)
      this.getAllMildstones(nextProps.projectDocument);
  }

  private getAllMildstones(projectDocument: string) {
    sp.web.lists.getByTitle(projectDocument).items
      .select("File", "Author/ID", "Author/Title").expand("File", "Author")
      .get()
      .then((response: Array<Document>) => {
        response.forEach(item => {
          item.FileName = item.File ? item.File.Name : "";
          item.OwnerTitle = item.Author ? item.Author.Title : "";
          item.Date = new Date(item.Created);
        });
        this.setState({ documentList: response })
      });

  }

  private ownerTemplate(rowData: Document, column) {
    if (rowData.Author) {
      return (
        <div className="taskListTitle"><span title={rowData.Author.Title}>{ rowData.Author.Title}</span></div>
      );
    }
  }

  private documentTypeTemplate(rowData: Document, column) {
    let type = "";
    let iconClass = "";
    if (rowData.File) {
      let data = rowData.File.Name.split(".");
      if (data.length > 1) {
        type = data[1];
      }
      switch (type.toLowerCase()) {
        case "doc":
        case "docx":
          iconClass = "far fa-file-word";
          break;
        case "pdf":
          iconClass = "far fa-file-pdf";
          break;
        case "xls":
        case "xlsx":
          iconClass = "far fa-file-excel";
          break;
        case "png":
        case "jpeg":
        case "gif":
          iconClass = "far fa-file-image";
          break;
        default:
          iconClass = "far fa-file";
          break;
      }
      return (
        <div className="taskListTitle">
          <i
            className={iconClass}
            style={{ marginRight: "5px" }}
          />
          <a href={rowData.File.ServerRelativeUrl} target="_blank" title={rowData.File.Name}>
            {rowData.File.Name}
          </a>
        </div>
        // <span>{rowData.File.Name}</span>
      );
    }
  }

  private timeDateTemplate(rowData: Document, column) {
    return (
      <span>{moment(rowData.Created).format("DD MMM YYYY")}</span>
    );
  }

  public render(): React.ReactElement<IProjectDocumentProps> {
    return (
      <div className="col-xs-12 col-md-5 cardPadding">
        <div className="well recommendedProjects">
          <div className="row">
            <div className="col-sm-12 col-12 cardHeading">
              <h5>Project Documents</h5>
            </div>
            <div className="col-sm-12 col-12 profileDetails-container">
              <DataTable
                value={this.state.documentList}
                responsive={true}
              >
                <Column
                  field="FileName"
                  header="Documents Type"
                  body={this.documentTypeTemplate}
                  sortable={true}
                  style={{ width: "50%" }}
                />
                <Column
                  field="OwnerTitle"
                  header="Owner"
                  body={this.ownerTemplate}
                  sortable={true}
                  style={{ width: "25%" }}
                />
                <Column
                  field="Date"
                  header="Date"
                  sortable={true}
                  body={this.timeDateTemplate}
                  style={{ width: "25%" }}
                />
              </DataTable>
            </div>
          </div>
        </div>
      </div>
    )
  }
}
