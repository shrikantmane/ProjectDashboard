import * as React from "react";
import { ICommentsProps } from './ICommentProps';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import styles from "../ProjectManagement.module.scss";
import { sp, ItemAddResult, Web } from "@pnp/sp";
import moment from 'moment/src/moment';

export default class Comments extends React.Component<ICommentsProps, {
    showPanel: boolean;
    commentList: any;
    fields: {},
    showButton: boolean
}>{
    constructor(props) {
        super(props);
        this.state = {
            showPanel: true,
            commentList: [],
            fields: {},
            showButton: false
        }
        this.sendComment = this.sendComment.bind(this);
        this.handleKeyUp = this.handleKeyUp.bind(this);

        console.log('id', this.props.id);
        console.log('list', this.props.list);
    }
    componentDidMount() {
        this.getTaskComments();
    }
    private getTaskComments() {
        let reactHandler = this;
        //let ScheduleList = currentProject.Schedule_x0020_List;
        let CommentList = this.props.list;
        sp.web.lists
            .getByTitle(CommentList)
            .items.select(
            "Created",
            "Comment",
            "Author/ID",
            "Author/Title"
            )
            .expand("Author")
            .filter("Task_x0020_Name/ID eq " + this.props.id)
            .get()
            .then(response => {
                console.log("response of Task List -", response);
                this.setState({ commentList: response });
                console.log('state:', this.state.commentList);
            });
    }

    private _closePanel = (): void => {
        this.setState({ showPanel: false });
        this.props.parentReopen();
    };
    handleKeyUp() {
        console.log(this.state.fields['text']);
        if (this.state.fields['text'] && this.state.fields['text'].trim()) {
            this.setState({ showButton: true });
        } else {
            this.setState({ showButton: false });
        }
    }
    handleChange(field, e) {
        let fields = this.state.fields;
        fields[field] = e.target.value;
        this.setState({ fields });
    }
    sendComment() {
        let ScheduleList = this.props.list;
        sp.web.lists.getByTitle(ScheduleList).items.add({
            Comment: this.state.fields['text'].trim(),
            Task_x0020_NameId: this.props.id
        }).then(r => {
            // this will add an attachment to the item we just created
            // r.item.attachmentFiles.add("file.txt", "Here is some file content.");
            this.getTaskComments();
            let field = this.state.fields;
            field['text'] = "";
            this.setState({ fields: field, showButton: false })
        });
    }
    public render(): React.ReactElement<ICommentsProps> {
        const commentContainer = this.state.commentList.length > 0 ?
            this.state.commentList.map(function (element, index) {
                return < li className="message left appeared" >
                    <div className="text_wrapper">
                        <div className="text">{element.Comment}</div>
                    </div>
                    <div className="name_date"><p>{element.Author.Title}</p>,<p> {moment(element.Created).format("DD MMM YYYY")} |</p><p>{moment(element.Created).format("hh:mm:ss a")}</p></div>
                </li >
            })
            : <li className="message left appeared"><div> No comments Found </div></li>
        const buttonContainer = this.state.showButton ?
            <div className="send_message">
                <div className="text" onClick={this.sendComment}>
                    <i className="fas fa-paper-plane"></i>
                </div>
            </div> : null;
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
                            <ul className="messages">
                                {commentContainer}
                            </ul>
                             <div className="bottom_wrapper clearfix">
                                <div className="message_input_wrapper">
                                    <input className="message_input" placeholder="Type your comment here..." onChange={this.handleChange.bind(this, "text")} onKeyUp={this.handleKeyUp}
                                        value={this.state.fields["text"]} />
                                </div>
                                {buttonContainer}
                            </div> 
                        </div>
                    </div>
                </Panel>
            </div>
        )
    }
}