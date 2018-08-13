import * as React from "react";
import { ICommentsProps } from './ICommentProps';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import styles from "../../ProjectDashboard.module.scss";
import { sp, ItemAddResult, Web } from "@pnp/sp";
import moment from 'moment/src/moment';

export default class Comments extends React.Component<ICommentsProps, {
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
    public render(): React.ReactElement<ICommentsProps> {

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
                                <div className="title">Comments</div>
                            </div>
                            <ul className="messages"><li className="message left appeared">
                                <div className="text_wrapper">
                                    <div className="text">Hello Alpha Server Team! :)</div>
                                </div>
                                <div className="name_date"><p>Chandler Bing</p>,<p> 8/8/2018 |</p><p>12:45 PM</p></div>
                            </li><li className="message left appeared">
                                    <div className="text_wrapper">
                                        <div className="text">A message is a discrete unit of communication intended by the source for consumption by some recipient or group of recipients. A message may be delivered by various means, including courier, telegraphy, carrier pigeon and electronic bus. A message can be the content of a broadcast.</div>
                                    </div>
                                    <div className="name_date"><p>Chandler Bing</p>,<p> 8/8/2018 |</p><p>12:45 PM</p></div>
                                </li><li className="message left appeared">
                                    <div className="text_wrapper">
                                        <div className="text">Going Good!</div>
                                    </div>
                                    <div className="name_date"><p>Chandler Bing</p>,<p> 8/8/2018 |</p><p>12:45 PM</p></div>
                                </li></ul>
                            <div className="bottom_wrapper clearfix">
                                <div className="message_input_wrapper">
                                    <input className="message_input" placeholder="Type your comment here..." />
                                </div>
                                <div className="send_message">
                                    <div className="text">
                                        <i className="fas fa-paper-plane"></i>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div className="message_template">
                            <li className="message">
                                <div className="text_wrapper">
                                    <div className="text"></div>
                                </div>
                                <div className="name_date">
                                    <p>Chandler Bing</p>,
																											<p> 8/8/2018 |</p>
                                    <p>12:45 PM</p>
                                </div>
                            </li>
                        </div>
                    </div>
                </Panel>
            </div>
        )
    }
}