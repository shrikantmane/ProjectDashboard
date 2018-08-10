import * as React from "react";
import { IDocumentsProps } from './IDocumentProps';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import styles from "../../ProjectDashboard.module.scss";
import { sp, ItemAddResult, Web } from "@pnp/sp";
import moment from 'moment/src/moment';

export default class Documents extends React.Component<IDocumentsProps, {
    showPanel: boolean;
    attachment: any;
    fields: {},
    errors: {},
}>{
    constructor(props) {
        super(props);
        this.state = {
            showPanel: true,
            attachment: [],
            fields: {},
            errors: {},
        }
        this.onSubmit = this.onSubmit.bind(this);
    }
    componentDidMount() {
        this.getTaskDocuments();
    }
    private getTaskDocuments() {
        let reactHandler = this;
        //let ScheduleList = currentProject.Schedule_x0020_List;
        let ScheduleList = 'AlphaServe_Schedule_List';
        sp.web.lists
            .getByTitle(ScheduleList)
            .items.select(
            "Created",
            "Attachments",
            "AttachmentFiles",
            "AttachmentFiles/ServerRelativeUrl",
            "AttachmentFiles/FileName",
            "Author/ID",
            "Author/Title"
            )
            .expand("AttachmentFiles", "Author")
            .filter("ID eq 10")
            .get()
            .then(response => {
                console.log("response of Schedule List -", response);
                if (response.length > 0) {
                    if (response[0].AttachmentFiles.length > 0) {
                        response[0].AttachmentFiles.forEach(element => {
                            element.Created = response[0].Created;
                            element.Auther = response[0].Author.Title;
                        });
                    }
                    this.setState({ attachment: response[0].AttachmentFiles });
                }
                console.log('State', this.state.attachment);
                //reactHandler.addAttachment(ScheduleList );
            });

    }

    private _closePanel = (): void => {
        this.setState({ showPanel: false });
        this.props.parentReopen();
    };
    handleChange(field, e) {
        let fields = this.state.fields;
        fields[field] = e.target.files[0];
        this.setState({ fields });
    }
    handleValidation() {
        let fields = this.state.fields;
        let errors = {};
        let errorClass = {};
        let formIsValid = true;
        if (!fields["attachment"]) {
            formIsValid = false;
            errors["attachment"] = "Please select attachment";
            errorClass["attachment"] = "classError";
        }
        this.setState({ errors: errors });
        return formIsValid;
    }
    onSubmit() {
        if (this.handleValidation()) {
            let ScheduleList = 'AlphaServe_Schedule_List';

            sp.web.lists.getByTitle(ScheduleList).items.getById(10).update({
                // Title: "My New Title",
            }).then(i => {
                i.item.attachmentFiles.add(this.state.fields["attachment"].name, this.state.fields["attachment"]).then((result) => {
                    this.getTaskDocuments();
                })
            });

        }
    }
    public render(): React.ReactElement<IDocumentsProps> {
        const fileContainer = this.state.attachment.length > 0 ?
            this.state.attachment.map(function (element, index) {
                let iconClass = "";
                let type = "";
                let data = element.FileName.split(".");
                if (data.length > 1) {
                    type = data[1];
                }
                switch (type.toLowerCase()) {
                    case "doc":
                    case "docx":
                        iconClass = "fas fa-file-word";
                        break;
                    case "pdf":
                        iconClass = "fas fa-file-pdf";
                        break;
                    case "xls":
                    case "xlsx":
                        iconClass = "fas fa-file-excel";
                        break;
                    case "png":
                    case "jpeg":
                    case "gif":
                        iconClass = "fas fa-file-image";
                        break;
                    default:
                        iconClass = "fas fa-file";
                        break;
                }
                return <div className="document-container">
                    <div className="ownerName_time">
                        <p>{element.Auther}</p>,
						<p>{moment(element.Created).format("DD MMM YYYY")} |</p>
                        <p>{moment(element.Created).format("hh:mm:ss a")}</p>
                    </div>
                    <div className="documents">
                        <p>
                            <i className={iconClass + " fa-2x"}></i>
                        </p>
                        <p>{element.FileName}</p>
                    </div>
                </div>
            })
            : <div className="document-container">No Files Found</div>
        return (
            <div>
                <Panel
                    isOpen={this.state.showPanel}
                    onDismiss={this._closePanel}
                    type={PanelType.medium}

                >
                    <div className="DocumentContainer">
                        <div className="chat_window">
                            <div className="chat_header">
                                <div className="title">Documents</div>
                            </div>
                            <div className="mid-section">
                                {fileContainer}
                            </div>
                            <div className="bottom_wrapper clearfix">
                                <div className="display-line form-group">
                                    <div className=" form-control fileupload" data-provides="fileupload">
                                        <input id="file_upload" type="file" onChange={this.handleChange.bind(this, "attachment")} />
                                        <span className="fileupload-new">Select file</span>
                                        <span className="fileupload-preview"></span>
                                    </div>
                                    <div style={{ marginLeft: '12px' }}>
                                        <button type="button" className="btn-style btn btn-primary" onClick={this.onSubmit}>Submit</button>
                                    </div>
                                    <span className="error">{this.state.errors["attachment"]}</span>
                                </div>
                            </div>
                        </div>
                    </div>
                </Panel>
            </div>
        )
    }
}