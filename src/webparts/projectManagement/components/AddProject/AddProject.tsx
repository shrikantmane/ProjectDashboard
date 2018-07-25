import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IAddProjectProps } from './IAddProjectProps';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import "bootstrap/dist/css/bootstrap.min.css";
import ProjectListTable from '../ProjectList/ProjectListTable';
import { Button, Modal } from 'react-bootstrap';


export default class AddProject extends React.Component<IAddProjectProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass:{},
    cloneProjectChecked: boolean,
    showModal: boolean;
}> {

    constructor(props) {
        super(props);
        this.state = {
            showPanel: true,
            fields: {},
            errors: {},
            errorClass:{},
            cloneProjectChecked: false,
            showModal: false
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
        } else {
            let fields = this.state.fields;
            fields[field] = e.target.value;
            this.setState({ fields });
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
        if (!fields["ownername"]) {
            formIsValid = false;
            errors["ownername"] = "Cannot be empty";
            errorClass["ownername"] = "classError";
        }
        if (!fields["projectdescription"]) {
            formIsValid = false;
            errors["projectdescription"] = "Cannot be empty";
            errorClass["projectdescription"] = "classError";
        }
        if (!fields["startdate"]) {
            formIsValid = false;
            errors["startdate"] = "Cannot be empty";
            errorClass["startdate"] = "classError";
        }
        if (!fields["duedate"]) {
            formIsValid = false;
            errors["duedate"] = "Cannot be empty";
            errorClass["duedate"] = "classError";
        }
        if (!fields["tags"]) {
            formIsValid = false;
            errors["tags"] = "Cannot be empty";
            errorClass["tags"] = "classError";
        }

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

    projectSubmit(e) {
        e.preventDefault();
        if (this.handleValidation()) {
            let obj: any = this.state.fields;
            sp.web.lists.getByTitle("Project").items.add({
                Project: obj.projectname ? obj.projectname : '',
                StartDate: obj.startdate ? new Date(obj.startdate) : '',
                DueDate: obj.duedate ? new Date(obj.duedate) : '',
                //AssignedToId: { results: [8] },
                Priority: obj.priority ? obj.priority : '',
                Body: obj.projectdescription ? obj.projectdescription : '',
                Department_x0020_Specific: obj.departmentspecific ? (obj.departmentspecific === 'on' ? true : false) : null,
                Recurring_x0020_Project: obj.requringproject ? (obj.requringproject === 'on' ? true : false) : null,
                Occurance: obj.occurance ? obj.occurance : '',
                //DepartmentId: 2,
                //Status0Id: 2

            }).then((response) => {
                console.log('Item adding-', response);
                this._closePanel();
                this._showModal();
                this.props.parentMethod();
            });
        } else {
            console.log("Form has errors.")
        }
    }
    _showModal() {
        this.setState({ showModal: true });
    };
    _closeModal() {
        this.setState({ showModal: false });
    };

    public render(): React.ReactElement<IAddProjectProps> {
        let formControl = 'form-control';
        let paddingInputStyle = 'padding-input-style';
        const selectProjectContent = this.state.cloneProjectChecked ?
            <div className="col-lg-12">
                <div className="form-group">
                    <label>Select Project</label>
                    <select className="form-control" ref="project" onChange={this.handleChange.bind(this, "project")} value={this.state.fields["project"]}>
                        <option>Project 1</option>
                        <option>Project 2</option>
                        <option>Project 3</option>
                    </select>
                </div>
            </div> : null;

        const chechbox1Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Schedule" onChange={this.handleChange.bind(this, "cloneschedule")} value={this.state.fields["cloneschedule"]} />
                    </div>
                </div>
            </div> : null;
        const chechbox2Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Documents" onChange={this.handleChange.bind(this, "clonedocuments")} value={this.state.fields["clonedocuments"]} />
                    </div>
                </div>
            </div> : null;
        const chechbox3Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Requirements" onChange={this.handleChange.bind(this, "clonerequirements")} value={this.state.fields["clonerequirements"]} />
                    </div>
                </div>
            </div> : null;
        const chechbox4Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Calender" onChange={this.handleChange.bind(this, "clonecalender")} value={this.state.fields["clonecalender"]} />
                    </div>
                </div>
            </div> : null;
        return (
            // className="PanelContainer"
            <div>
                <Panel
                    isOpen={this.state.showPanel}
                    onDismiss={this._closePanel}
                    type={PanelType.medium}

                >
                    <div className="PanelContainer">
                        <section className="main-content-section">

                            <div className="wrapper">

                                <div className="row">

                                    <div className="col-md-12">
                                        <section id="step1">
                                            <div className="well">
                                                <div className="row">
                                                    <h3>Project Details</h3>
                                                    <div >
                                                        <form name="projectform" onSubmit={this.projectSubmit.bind(this)}>
                                                            <div className="row">
                                                                <div className="col-lg-12">
                                                                    <div className="form-group">
                                                                        <label>Clone Project</label>
                                                                        <div>
                                                                            <Checkbox onChange={this.handleChange.bind(this, "cloneproject")} value={this.state.fields["cloneproject"]} />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                {selectProjectContent}
                                                                {chechbox1Content}
                                                                {chechbox2Content}
                                                                {chechbox3Content}
                                                                {chechbox4Content}

                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Project Name</label>
                                                                        <input ref="projectname" type="text" className={formControl + " " + (this.state.errorClass["projectname"] ? this.state.errorClass["projectname"] : '')} placeholder="Enter project name"
                                                                            onChange={this.handleChange.bind(this, "projectname")} value={this.state.fields["projectname"]}>
                                                                        </input>
                                                                        <span className="error">{this.state.errors["projectname"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Owner</label>
                                                                        <span className="calendar-style"><i className="fas fa-user icon-style"></i>
                                                                            <input ref="ownername" type="text" className={paddingInputStyle + " " + formControl + " " + (this.state.errorClass["ownername"] ? this.state.errorClass["ownername"] : '')} placeholder="Enter owners name"
                                                                                onChange={this.handleChange.bind(this, "ownername")} value={this.state.fields["ownername"]}>
                                                                            </input>
                                                                        </span>
                                                                        <span className="error">{this.state.errors["ownername"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-12">
                                                                    <div className="form-group">
                                                                        <label>Project Description</label>
                                                                        <textarea ref="projectdescription" className={formControl + " " + (this.state.errorClass["projectdescription"] ? this.state.errorClass["projectdescription"] : '')} placeholder="Brief the owner about the project"
                                                                            onChange={this.handleChange.bind(this, "projectdescription")} value={this.state.fields["projectdescription"]}></textarea>
                                                                        <span className="error">{this.state.errors["projectdescription"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Start Date</label>
                                                                        <DatePicker
                                                                            placeholder="Select start date"
                                                                            onSelectDate={this.handleChange.bind(this, "startdate")}
                                                                            value={this.state.fields["startdate"]}
                                                                        />
                                                                        <span className="error">{this.state.errors["startdate"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Due Date</label>
                                                                        <DatePicker
                                                                            placeholder="Select due date"
                                                                            onSelectDate={this.handleChange.bind(this, "duedate")}
                                                                            value={this.state.fields["duedate"]}
                                                                        />
                                                                        <span className="error">{this.state.errors["duedate"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Priority</label>
                                                                        <select className={formControl + " " + (this.state.errorClass["priority"] ? this.state.errorClass["priority"] : '')} ref="priority" onChange={this.handleChange.bind(this, "priority")} value={this.state.fields["priority"]}>
                                                                            <option>Low</option>
                                                                            <option>Medium</option>
                                                                            <option>High</option>
                                                                        </select>
                                                                        <span className="error">{this.state.errors["priority"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Project Type</label>
                                                                        <div>
                                                                            <Checkbox label="Department Specific" onChange={this.handleChange.bind(this, "departmentspecific")} value={this.state.fields["departmentspecific"]} />
                                                                        </div>
                                                                    </div>
                                                                </div>

                                                                <div className="clearfix"></div>

                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Tags</label>
                                                                        <input ref="tags" type="text" className={formControl + " " + (this.state.errorClass["tags"] ? this.state.errorClass["tags"] : '')} placeholder="Enter Tags"
                                                                            onChange={this.handleChange.bind(this, "tags")} value={this.state.fields["tags"]}>
                                                                        </input>
                                                                        <span className="error">{this.state.errors["tags"]}</span>
                                                                    </div>
                                                                </div>

                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Requring Project?</label>
                                                                        <div>
                                                                            <Checkbox label="Yes" onChange={this.handleChange.bind(this, "requringproject")} value={this.state.fields["requringproject"]} />
                                                                        </div>
                                                                        {/* <div className="display-line">
                                                                            <span className="col-lg-12 col-sm-12 radBtn">
                                                                                <input ref="requringproject" type="checkbox" id="2" name="selectorAssignor"
                                                                                    onChange={this.handleChange.bind(this, "requringproject")} value={this.state.fields["requringproject"]}>
                                                                                </input>
                                                                                <div className="check"></div>
                                                                                <p className="checkbox-title">Yes	</p>
                                                                            </span>
                                                                        </div> */}
                                                                    </div>
                                                                </div>
                                                                <div className="clearfix"></div>
                                                                <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Occurance</label>
                                                                        <select ref="occurance" className={formControl + " " + (this.state.errorClass["occurance"] ? this.state.errorClass["occurance"] : '')}
                                                                            onChange={this.handleChange.bind(this, "occurance")} value={this.state.fields["occurance"]}>
                                                                            <option>Daily</option>
                                                                            <option>Weekly </option>
                                                                            <option>Months</option>
                                                                        </select>
                                                                        <span className="error">{this.state.errors["occurance"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-lg-12">
                                                                    <div className="btn-sec">
                                                                        <button id="submit" value="Submit" className="btn-style btn btn-success">Save</button>
                                                                        <button type="button" className="btn-style btn btn-default" onClick={this._closePanel}>Cancel</button>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </form>
                                                    </div>
                                                </div>
                                            </div>

                                        </section>
                                    </div>
                                </div>

                            </div>
                        </section>
                    </div>
                </Panel>

                <Modal
                    show={this.state.showModal}
                    onHide={this._closeModal}
                    container={this}
                    aria-labelledby="contained-modal-title"
                    animation={false}
                >
                    <Modal.Header>
                        <Modal.Title id="contained-modal-title">
                            Project Created
                        </Modal.Title>
                    </Modal.Header>
                    <Modal.Body>
                        Project Created Successfully! Do you want to configure Project
Schedule and Project Team now?
                    </Modal.Body>
                    <Modal.Footer>
                        <Button onClick={this._closeModal}>I'll Do it Later</Button>
                        <Button onClick={this._closeModal}>Continue</Button>
                    </Modal.Footer>
                </Modal>
            </div>

        );
    }
    private _closePanel = (): void => {
        this.setState({ showPanel: false });
    };
    /* Api Call*/

}
