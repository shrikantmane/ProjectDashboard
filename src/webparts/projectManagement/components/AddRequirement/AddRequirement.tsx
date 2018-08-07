import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IAddRequirementProps } from './IAddRequirementProps';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import "bootstrap/dist/css/bootstrap.min.css";
import { Button, Modal } from 'react-bootstrap';
import ProjectListTable from '../ProjectList/ProjectListTable';

const textcolor = {
    color: 'red' as 'red',
}

export default class AddProject extends React.Component<IAddRequirementProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass: {},
    cloneProjectChecked: boolean,
    showModal: boolean,
    isDataSaved: boolean,
    attachmentFiles: any,
   // isLoading: boolean
}> {

    constructor(props) {
        super(props);
        this.state = {
            showPanel: true,
            fields: {},
            errors: {},
            errorClass: {},
            cloneProjectChecked: false,
            showModal: false,
            isDataSaved: false,
            attachmentFiles: [],
           // isLoading: false
        };
        this._showModal = this._showModal.bind(this);
        this._closeModal = this._closeModal.bind(this);
    }

    handleChange(field, e) {
        if (field === 'startdate') {
            let fields = this.state.fields;
            fields[field] = e;
        }
        else if (field === 'duedate') {
            let fields = this.state.fields;
            fields[field] = e;
        }
        else if (field === 'cloneproject') {
            let fields = this.state.fields;
            fields[field] = e.target.value;
            this.setState({ fields, cloneProjectChecked: !this.state.cloneProjectChecked });
        } else if (field === 'filedescription') {
            let fields = this.state.fields;
            fields[field] = e.target.files[0]
            this.setState({ fields });
        } else {
            let fields = this.state.fields;
            fields[field] = e.target.value;
            this.setState({ fields });
        }
    }
    componentDidMount() {
        if (this.props.id) {
            this.getProjectByID(this.props.id);
            this.setState({
                fields: {}
            })

        } else {
            this.setState({
                fields: {}
            })
        }
    }

    handleValidation() {
        let fields = this.state.fields;
        let errors = {};
        let errorClass = {};
        let formIsValid = true;

        //Name
        if (!fields["projectname"]) {
            formIsValid = false;
            errors["projectname"] = "Cannot be empty";
            errorClass["projectname"] = "classError";
        }
        // if (!fields["projectdescription"]) {
        //     formIsValid = false;
        //     errors["projectdescription"] = "Cannot be empty";
        //     errorClass["projectdescription"] = "classError";
        // }
        // if (!fields["effortdescription"]) {
        //     formIsValid = false;
        //     errors["effortdescription"] = "Cannot be empty";
        //     errorClass["effortdescription"] = "classError";
        // }
        // if (!fields["filedescription"]) {
        //     formIsValid = false;
        //     errors["filedescription"] = "Cannot be empty";
        //     errorClass["filedescription"] = "classError";
        // }



        // if (typeof fields["name"] !== "undefined") {
        //     if (!fields["name"].match(/^[a-zA-Z]+$/)) {
        //         formIsValid = false;
        //         errors["name"] = "Only letters";
        //     }
        // }

        //Email
        // if (!fields["email"]) {
        //     formIsValid = false;
        //     errors["email"] = "Cannot be empty";
        // }

        // if (typeof fields["email"] !== "undefined") {
        //     let lastAtPos = fields["email"].lastIndexOf('@');
        //     let lastDotPos = fields["email"].lastIndexOf('.');

        //     if (!(lastAtPos < lastDotPos && lastAtPos > 0 && fields["email"].indexOf('@@') == -1 && lastDotPos > 2 && (fields["email"].length - lastDotPos) > 2)) {
        //         formIsValid = false;
        //         errors["email"] = "Email is not valid";
        //     }

        this.setState({ errors: errors, errorClass: errorClass });
        return formIsValid;
    }
    private getProjectByID(id): void {
        // get Project Documents list items for all projects
        let filterString = "ID eq " + id;
        sp.web.lists.getByTitle(this.props.list).items
            .select("ID", "Requirement", "Resources", "Impact_x0020_On_x0020_Timelines", "Efforts", "Attachments", "Apporval_x0020_Status", "Approver/Title", "Approver/ID", "Author/Title", "Author/ID", "Created", "AttachmentFiles", "AttachmentFiles/ServerRelativeUrl", "AttachmentFiles/FileName")
            .expand("Approver", "Author", "AttachmentFiles")
            .filter(filterString)
            .get()
            .then((response) => {
                let fields = this.state.fields;
                console.log('Project1 by name', response);
                console.log('Project112 by name', response[0].Requirement);
                fields["projectname"] = response ? response[0].Requirement : '';
                fields["projectdescription"] = response ? response[0].Resources : '';
                fields["effortdescription"] = response ? response[0].Efforts : '';

                this.setState({ fields, attachmentFiles: response[0].AttachmentFiles });
                //    this.setState({
                //     fields: response[0].Roles_Responsibility
                //    })
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });

    }

    projectSubmit(e) {
        e.preventDefault();
        if (this.handleValidation()) {

            let obj: any = this.state.fields;
            if (this.props.id) {
               // this.setState({ isLoading: true });
                sp.web.lists.getByTitle(this.props.list).items.getById(this.props.id).update({
                    Requirement: obj.projectname ? obj.projectname : '',
                    //Target_x0020_Date: obj.startdate ? new Date(obj.startdate) : '',
                    Resources: obj.projectdescription ? obj.projectdescription : null,
                    Efforts: obj.effortdescription ? obj.effortdescription : null,
                    // Attachments: obj.filedescription ? obj.effortdescription : '',
                    //Owner: obj.ownername?obj.ownername:'',
                    // Impact: obj.priority ? obj.priority : '',
                    // Mitigation: obj.projectdescription ? obj.projectdescription : '',
                    //Department_x0020_Specific: obj.departmentspecific ? (obj.departmentspecific === 'on' ? true : false) : null,
                    //Recurring_x0020_Project: obj.requringproject ? (obj.requringproject === 'on' ? true : false) : null,
                    //Occurance: obj.occurance ? obj.occurance : '',
                    //DepartmentId: 2,
                    //Status0Id: 2

                }).then((response) => {
                    if(this.state.fields["filedescription"]!=null)
                    {
                    //response.item.attachmentFiles.add(this.state.fields["filedescription"], this.state.selectedFile);
                    response.item.attachmentFiles.add(this.state.fields["filedescription"].name, this.state.fields["filedescription"]).then((response) => {
                       // this.setState({ isLoading: false });
                        this._closePanel();
                        this.props.parentMethod();
                        this.props.parentReopen();
                    })
                }
                else{
                    this._closePanel();
                        this.props.parentMethod();
                        this.props.parentReopen();
                }
                });
            } else {
             //   this.setState({ isLoading: true });
                sp.web.lists.getByTitle(this.props.list).items.add({
                    Requirement: obj.projectname ? obj.projectname : '',
                    //Target_x0020_Date: obj.startdate ? new Date(obj.startdate) : '',
                    Resources: obj.projectdescription ? obj.projectdescription : null,
                    Efforts: obj.effortdescription ? obj.effortdescription : null,
                }).then((response) => {
                    if(this.state.fields["filedescription"]!=null)
                    {
                    //const formData = new FormData();
                    //formData.append('myFile', this.state.selectedFile, this.state.selectedFile.name);
                    response.item.attachmentFiles.add(this.state.fields["filedescription"].name, this.state.fields["filedescription"]).then((response) => {
                        //this.setState({ isLoading: false });
                        this.props.parentMethod();
                        this._closePanel();
                    })
                }
                else{
                    this.props.parentMethod();
                    this._closePanel();
                }
                    console.log('Item adding-', response);
                    this.setState({ isDataSaved: true });
                    //this._closePanel();
                    this._showModal();
                    //this.props.parentMethod();
                });
            }
        }
    }
    _showModal() {
        this.setState({ showModal: true });
    };
    _closeModal() {
        this.setState({ showModal: false });
    };
    removeAttachment(i, event) {
        var result = confirm("Are you sure you want to delete item?");
        if (result) {
        console.log('index1', i);
        let tempAttachment = this.state.attachmentFiles;
        tempAttachment.splice(i, 1);
        this.setState({ attachmentFiles: tempAttachment });
    }
}
    public render(): React.ReactElement<IAddRequirementProps> {
        let formControl = 'form-control';
        let paddingInputStyle = 'padding-input-style';
        const attachmentDiv = this.state.attachmentFiles.length > 0 ?
            <div className="col-lg-6">
                {this.state.attachmentFiles.map((obj, i) =>
                    <div className="form-group">
                        <label style={{ float: 'left', width: '90%' }}><a href={obj.ServerRelativeUrl}><i
                            style={{ marginRight: "5px" }}
                            className='fa fa-file' ></i>{obj.FileName}</a></label>
                        <i className="far fa-times-circle" style={{ float: 'right', cursor: 'pointer' }} onClick={this.removeAttachment.bind(this, i)}></i>
                    </div>
                )}
            </div> : null;
        const emptyDiv = this.state.attachmentFiles.length > 0 ?
            <div className="col-lg-6">
            </div> : null;
        // const chechbox1Content = this.state.cloneProjectChecked ?
        //     <div className="col-lg-6">
        //         <div className="form-group">
        //             <div>
        //                 <Checkbox label="Clone Schedule" onChange={this.handleChange.bind(this, "cloneschedule")} value={this.state.fields["cloneschedule"]} />
        //             </div>
        //         </div>
        //     </div> : null;
        // const chechbox2Content = this.state.cloneProjectChecked ?
        //     <div className="col-lg-6">
        //         <div className="form-group">
        //             <div>
        //                 <Checkbox label="Clone Documents" onChange={this.handleChange.bind(this, "clonedocuments")} value={this.state.fields["clonedocuments"]} />
        //             </div>
        //         </div>
        //     </div> : null;
        // const chechbox3Content = this.state.cloneProjectChecked ?
        //     <div className="col-lg-6">
        //         <div className="form-group">
        //             <div>
        //                 <Checkbox label="Clone Requirements" onChange={this.handleChange.bind(this, "clonerequirements")} value={this.state.fields["clonerequirements"]} />
        //             </div>
        //         </div>
        //     </div> : null;
        // const chechbox4Content = this.state.cloneProjectChecked ?
        //     <div className="col-lg-6">
        //         <div className="form-group">
        //             <div>
        //                 <Checkbox label="Clone Calender" onChange={this.handleChange.bind(this, "clonecalender")} value={this.state.fields["clonecalender"]} />
        //             </div>
        //         </div>
        //     </div> : null;
        return (
            // className="PanelContainer"
           // !this.state.isLoading ?
            <div>
                <Panel
                    isOpen={this.state.showPanel}
                    onDismiss={this._closePanel}
                    type={PanelType.medium}

                >
                    <div className="PanelContainer">
                        <section className="main-content-section">
                            {/* <div className="wrapper"> */}
                                <div className="row">
                                    <div className="col-sm-12 col-12">
                                        {/* <section id="step1">
                                            <div className="well">
                                                <div className="row"> */}
                                                    <h3 className="hbc-form-header">Project Requirement</h3>
                                                    {/* <div > */}
                                                        <form name="projectform" className="hbc-form" onSubmit={this.projectSubmit.bind(this)}>
                                                            <div className="row addSection">
                                                                <div className="col-sm-12 col-12">
                                                                    <div className="form-group">
                                                                    <span className="error">* </span><label>Requirements</label>
                                                                        <textarea ref="projectname" className={formControl + " " + (this.state.errorClass["projectname"] ? this.state.errorClass["projectname"] : '')} placeholder="Brief the owner about the project"
                                                                            onChange={this.handleChange.bind(this, "projectname")} value={this.state.fields["projectname"]}>
                                                                        </textarea>
                                                                        <span className="error">{this.state.errors["projectname"]}</span>
                                                                    </div>
                                                                </div>
                                                              
                                                                <div className="col-sm-6 col-12">
                                                                    <div className="form-group">
                                                                        <label>Number Of Resources</label>
                                                                        <input ref="projectdescription" type="number" className={formControl + " " + (this.state.errorClass["projectdescription"] ? this.state.errorClass["projectdescription"] : '')} placeholder="Total Number Of People"
                                                                            onChange={this.handleChange.bind(this, "projectdescription")} value={this.state.fields["projectdescription"]}>
                                                                        </input>
                                                                        <span className="error">{this.state.errors["projectdescription"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-sm-6 col-12">
                                                                    <div className="form-group">
                                                                        <label>Efforts</label>
                                                                        <input ref="effortdescription" type="number" className={formControl + " " + (this.state.errorClass["effortdescription"] ? this.state.errorClass["projectdescription"] : '')} placeholder="Enter Number Of Day"
                                                                            onChange={this.handleChange.bind(this, "effortdescription")} value={this.state.fields["effortdescription"]}>
                                                                        </input>
                                                                        <span className="error">{this.state.errors["effortdescription"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-sm-6 col-12">
                                                                    <div className="form-group">
                                                                        <label>Attachments</label>
                                                                        <div className="form-control fileupload" data-provides="fileupload">
                                                                            <input ref="filedescription" type="file" id="uploadFile" className={(this.state.errorClass["filedescription"] ? this.state.errorClass["filedescription"] : '')}
                                                                                onChange={this.handleChange.bind(this, "filedescription")} >
                                                                            </input>
                                                                            <span className="error">{this.state.errors["filedescription"]}</span>
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                                {emptyDiv}
                                                                {attachmentDiv}
                                                                </div>
                                                                <div className="row addSection">
                                                                {/* <div className="col-sm-12 col-12">
                                                                </div> */}
                                                                {/* <div className="clearfix"></div>
                                                                <div className="clearfix"></div> */}
                                                                <div className="col-sm-12 col-12">
                                                                    <div className="btn-sec">
                                                                        <button id="submit" value="Submit" className="btn-style btn btn-success">{this.props.id ? 'Update' : 'Save'}</button>
                                                                        <button type="button" className="btn-style btn btn-default" onClick={this._closePanel}>Cancel</button>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </form>
                                                    {/* </div> */}
                                                {/* </div>
                                            </div>
                                        </section> */}
                                    </div>
                                </div>

                            {/* </div> */}
                        </section>
                    </div>
                </Panel>
            </div>
//: <div style={{ textAlign: "center", fontSize: "25px" }}><i className="fa fa-spinner fa-spin"></i></div>
        );
    }
    private _closePanel = (): void => {
        this.setState({ showPanel: false });
        if (!this.state.isDataSaved) {
            this.props.parentReopen();
        }
    };
    /* Api Call*/

}
