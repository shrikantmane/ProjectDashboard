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
import { Button, Modal } from 'react-bootstrap';

import ProjectListTable from '../ProjectList/ProjectListTable';


import { Project } from "../ProjectList/ProjectList";
import { IPersonaProps, Persona, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { BaseComponent, assign } from 'office-ui-fabric-react/lib/Utilities';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    IBasePicker,
    ListPeoplePicker,
    NormalPeoplePicker,
    ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.types';


const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: true,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    suggestionsContainerAriaLabel: 'Suggested contacts'
};
const limitedSearchAdditionalProps: IBasePickerSuggestionsProps = {
    searchForMoreText: 'Load all Results',
    resultsMaximumNumber: 10,
    searchingText: 'Searching...'
};
const limitedSearchSuggestionProps: IBasePickerSuggestionsProps = assign(limitedSearchAdditionalProps, suggestionProps);
export default class AddProject extends React.Component<IAddProjectProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass: {},
    cloneProjectChecked: boolean,
    showModal: boolean;
    projectList: any,
    peopleList: any[],
    delayResults: false,
    isDataSaved: boolean;
}> {
    private _picker: IBasePicker<IPersonaProps>;
    constructor(props) {
        super(props);
        const peopleList: IPersonaWithMenu[] = [];

        this.state = {
            showPanel: true,
            fields: {},
            errors: {},
            errorClass: {},
            cloneProjectChecked: false,
            showModal: false,
            projectList: new Array<Project>(),
            peopleList: peopleList,
            delayResults: false,
            isDataSaved: false
        };
        this._showModal = this._showModal.bind(this);
        this._closeModal = this._closeModal.bind(this);
    }
    componentDidMount() {
        this._getAllSiteUsers();
        if (this.props.id) {
            this.setState({
                fields: {}
            })
            this.getProjectByID(this.props.id);
            this.getProjectTagsByProjectName(this.props.id);
        } else {
            this.setState({
                fields: {}
            })
        }
    }
    componentWillReceiveProps(nextProps) {

    }
    private _getAllSiteUsers = (): void => {
        var reactHandler = this;
        sp.web.siteUsers.get().then(function (data) {
            const peopleList: IPersonaWithMenu[] = [];
            data.forEach((persona) => {
                const target: IPersonaWithMenu = {};
                let tempPersona = {
                    key: persona.Id,
                    text: persona.Title
                }
                assign(target, tempPersona);
                peopleList.push(target);

            });

            const mru: IPersonaProps[] = peopleList.slice(0, 5);
            reactHandler.setState({
                peopleList: peopleList,
                //mostRecentlyUsed: mru
            });
            console.log('People : ' + peopleList);
        });
    };

    handleChange(field, e, isChecked: boolean) {
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
            fields[field] = isChecked;
            this.setState({ fields, cloneProjectChecked: !this.state.cloneProjectChecked });
        }
        else if (field === 'cloneschedule') {
            let fields = this.state.fields;
            fields[field] = isChecked;
            this.setState({ fields });
        }
        else if (field === 'clonedocuments') {
            let fields = this.state.fields;
            fields[field] = isChecked;
            this.setState({ fields });
        }
        else if (field === 'clonerequirements') {
            let fields = this.state.fields;
            fields[field] = isChecked;
            this.setState({ fields });
        }
        else if (field === 'clonecalender') {
            let fields = this.state.fields;
            fields[field] = isChecked;
            this.setState({ fields });
        }
        else if (field === 'departmentspecific') {
            let fields = this.state.fields;
            fields[field] = isChecked;
            this.setState({ fields });
        }
        else if (field === 'requringproject') {
            let fields = this.state.fields;
            fields[field] = isChecked;
            this.setState({ fields });
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
        if (fields["startdate"] && fields["duedate"]) {
            if (fields["duedate"] < fields["startdate"]) {
                formIsValid = false;
                errors["duedate"] = "Due Date should always be greater than Start Date";
                errorClass["duedate"] = "classError";
            }
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
            if (this.props.id) {
                sp.web.lists.getByTitle("Project").items.getById(this.props.id).update({
                    StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
                    DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : '',
                    //Status0Id: 2,
                    //AssignedToId: 20,
                    Priority: obj.priority ? obj.priority : 'Low',
                    Clone_x0020_Project: obj.cloneproject ? obj.cloneproject : false,
                    Clone_x0020_Documents: obj.clonedocuments ? obj.clonedocuments : false,
                    Clone_x0020_Requirements: obj.clonerequirements ? obj.clonerequirements : false,
                    Clone_x0020_Schedule: obj.cloneschedule ? obj.cloneschedule : false,
                    Clone_x0020_Calender: obj.clonecalender ? obj.clonecalender : false,
                    Body: obj.projectdescription ? obj.projectdescription : '',
                    Occurance: obj.occurance ? obj.occurance : 'Daily',
                    Recurring_x0020_Project: obj.requringproject ? obj.requringproject : false,
                    ProTypeDeptSpecific: obj.departmentspecific ? obj.departmentspecific : false

                }).then(i => {
                    this._closePanel();
                    this.props.parentMethod();
                    this.props.parentReopen();
                });
            } else {
                sp.web.lists.getByTitle("Project").items.add({
                    Project: obj.projectname ? obj.projectname : '',
                    StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
                    DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : '',
                    //AssignedToId: { results: [8] },
                    Priority: obj.priority ? obj.priority : 'Low',
                    Body: obj.projectdescription ? obj.projectdescription : '',
                    ProTypeDeptSpecific: obj.departmentspecific ? obj.departmentspecific : false,
                    Recurring_x0020_Project: obj.requringproject ? obj.requringproject : false,
                    Occurance: obj.occurance ? obj.occurance : 'Daily',
                    Clone_x0020_Project: obj.cloneproject ? obj.cloneproject : false,
                    Clone_x0020_Documents: obj.clonedocuments ? obj.clonedocuments : false,
                    Clone_x0020_Requirements: obj.clonerequirements ? obj.clonerequirements : false,
                    Clone_x0020_Schedule: obj.cloneschedule ? obj.cloneschedule : false,
                    Clone_x0020_Calender: obj.clonecalender ? obj.clonecalender : false,
                    //DepartmentId: 2,
                    //Status0Id: 2

                }).then((response) => {
                    console.log('Item adding-', response);
                    this.setState({ isDataSaved: true });
                    this._closePanel();
                    this._showModal();

                });
            }
        } else {
            console.log("Form has errors.")
        }
    }
    _showModal() {
        this.setState({ showModal: true });
    };
    _closeModal() {
        this.setState({ showModal: false });
        this.props.parentMethod();
        this.props.parentReopen();
    };
    private _closePanel = (): void => {
        this.setState({ showPanel: false });
        if (!this.state.isDataSaved) {
             this.props.parentReopen();
        }
    };
    private _validateInput = (input: string): ValidationState => {
        if (input.indexOf('@') !== -1) {
            return ValidationState.valid;
        } else if (input.length > 1) {
            return ValidationState.warning;
        } else {
            return ValidationState.invalid;
        }
    };
    private _onInputChange(input: string): string {
        const outlookRegEx = /<.*>/g;
        const emailAddress = outlookRegEx.exec(input);

        if (emailAddress && emailAddress[0]) {
            return emailAddress[0].substring(1, emailAddress[0].length - 1);
        }

        return input;
    }
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
                        <Checkbox label="Clone Schedule" checked={this.state.fields["cloneschedule"]} onChange={this.handleChange.bind(this, "cloneschedule")} value={this.state.fields["cloneschedule"]} />
                    </div>
                </div>
            </div> : null;
        const chechbox2Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Documents" checked={this.state.fields["clonedocuments"]} onChange={this.handleChange.bind(this, "clonedocuments")} value={this.state.fields["clonedocuments"]} />
                    </div>
                </div>
            </div> : null;
        const chechbox3Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Requirements" checked={this.state.fields["clonerequirements"]} onChange={this.handleChange.bind(this, "clonerequirements")} value={this.state.fields["clonerequirements"]} />
                    </div>
                </div>
            </div> : null;
        const chechbox4Content = this.state.cloneProjectChecked ?
            <div className="col-lg-6">
                <div className="form-group">
                    <div>
                        <Checkbox label="Clone Calender" checked={this.state.fields["clonecalender"]} onChange={this.handleChange.bind(this, "clonecalender")} value={this.state.fields["clonecalender"]} />
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
                                                                            <Checkbox checked={this.state.fields["cloneproject"]} onChange={this.handleChange.bind(this, "cloneproject")} value={this.state.fields["cloneproject"]} />
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
                                                                            {/* <CompactPeoplePicker
                                                                            onResolveSuggestions={this._onFilterChangedWithLimit}
                                                                            getTextFromItem={this._getTextFromItem}
                                                                            className={'ms-PeoplePicker'}
                                                                            onGetMoreResults={this._onFilterChanged}
                                                                            pickerSuggestionsProps={limitedSearchSuggestionProps}
                                                                            inputProps={{
                                                                                onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called', ev),
                                                                                onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called', ev),
                                                                                'aria-label': 'People Picker'
                                                                            }}
                                                                            resolveDelay={300}
                                                                        /> */}
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
                                                                            <Checkbox checked={this.state.fields["departmentspecific"]} label="Department Specific" onChange={this.handleChange.bind(this, "departmentspecific")} value={this.state.fields["departmentspecific"]} />
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
                                                                            <Checkbox label="Yes" checked={this.state.fields["requringproject"]} onChange={this.handleChange.bind(this, "requringproject")} value={this.state.fields["requringproject"]} />
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
                                                                        <button id="submit" value="Submit" className="btn-style btn btn-success">{this.props.id ? 'Update' : 'Save'}</button>
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

    /* Api Call*/
    private getProjectByID(id): void {
        // get Project Documents list items for all projects
        let filterString = "ID eq " + id;
        sp.web.lists.getByTitle("Project").items
            .select("Project", "StartDate", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "Priority", "Clone_x0020_Project", "Clone_x0020_Calender", "Clone_x0020_Documents", "Clone_x0020_Requirements", "Clone_x0020_Schedule", "Body", "Occurance", "Recurring_x0020_Project", "ProTypeDeptSpecific")
            .expand("Status0", "AssignedTo")
            .filter(filterString)
            .getAll()
            .then((response) => {
                console.log('Project by name', response);
                this.state.fields["projectname"] = response ? response[0].Project : '';
                this.state.fields["priority"] = response ? response[0].Priority : '';
                this.state.fields["duedate"] = response ? new Date(response[0].DueDate) : '';
                this.state.fields["ownername"] = response ? response[0].AssignedTo[0].Title : '';
                this.state.fields["projectdescription"] = response ? response[0].Body : '';
                this.state.fields["startdate"] = response ? new Date(response[0].StartDate) : '';
                this.state.fields["departmentspecific"] = response ? response[0].ProTypeDeptSpecific : false;
                this.state.fields["requringproject"] = response ? response[0].Recurring_x0020_Project : false;
                this.state.fields["occurance"] = response ? response[0].Occurance : '';
                this.state.fields["cloneschedule"] = response ? response[0].Clone_x0020_Schedule : false;
                this.state.fields["clonedocuments"] = response ? response[0].Clone_x0020_Documents : false;
                this.state.fields["clonerequirements"] = response ? response[0].Clone_x0020_Requirements : false;
                this.state.fields["clonecalender"] = response ? response[0].Clone_x0020_Calender : false;
                this.state.fields["cloneproject"] = response ? response[0].Clone_x0020_Project : false;
                if (response[0].Clone_x0020_Project) {
                    this.setState({ cloneProjectChecked: true });
                } else {
                    this.setState({ cloneProjectChecked: false });
                }
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });
    }
    private getProjectTagsByProjectName(id): void {
        let filterString = "Project/ID eq " + id;
        //let filterString = "Project/ID eq 1";
        sp.web.lists.getByTitle("Project Tags").items
            .select("Project/ID", "Project/Title", "Tag").expand("Project")
            .filter(filterString)
            .get()
            .then((response) => {
                console.log('Project tag 1 -', response);
                this.state.fields["tags"] = response.length > 0 ? response[0].Tag : '';
            });
    }


    // PeoplePicker
    // private _onItemsChange = (items: any[]): void => {
    //     this.setState({
    //         selectedUser: items
    //     });
    // };

    // private _onItemSelected = (item: IPersonaProps): Promise<IPersonaProps> => {
    //     const processedItem = Object.assign({}, item);
    //     processedItem.text = `${item.text} (selected)`;
    //     return new Promise<IPersonaProps>((resolve, reject) => setTimeout(() => resolve(processedItem), 250));
    // };

    private _renderFooterText = (): JSX.Element => {
        return <div>No additional results</div>;
    };

    private _onFilterChangedWithLimit = (
        filterText: string,
        currentPersonas: IPersonaProps[]
    ): IPersonaProps[] | Promise<IPersonaProps[]> => {
        return this._onFilterChanged(filterText, currentPersonas, 3);
    };

    private _onFilterChanged = (
        filterText: string,
        currentPersonas: IPersonaProps[],
        limitResults?: number
    ): IPersonaProps[] | Promise<IPersonaProps[]> => {
        if (filterText) {
            let filteredPersonas: IPersonaProps[] = this._filterPersonasByText(filterText);

            filteredPersonas = this._removeDuplicates(filteredPersonas, currentPersonas);
            filteredPersonas = limitResults ? filteredPersonas.splice(0, limitResults) : filteredPersonas;
            return this._filterPromise(filteredPersonas);
        } else {
            return [];
        }
    };

    private _filterPersonasByText(filterText: string): IPersonaProps[] {
        return this.state.peopleList.filter(item => this._doesTextStartWith(item.text as string, filterText));
    }
    private _doesTextStartWith(text: string, filterText: string): boolean {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    }

    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }

    private _listContainsPersona(persona: any, personas: any) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.text === persona.text).length > 0;
    }

    private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        if (this.state.delayResults) {
            return this._convertResultsToPromise(personasToReturn);
        } else {
            return personasToReturn;
        }
    }

    private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }

    // private _returnMostRecentlyUsedWithLimit = (
    //     currentPersonas: IPersonaProps[]
    // ): IPersonaProps[] | Promise<IPersonaProps[]> => {
    //     let { mostRecentlyUsed } = this.state;
    //     mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
    //     mostRecentlyUsed = mostRecentlyUsed.splice(0, 3);
    //     return this._filterPromise(mostRecentlyUsed);
    // };

    private _getTextFromItem(persona: any): string {
        return persona.text as string;
    }

    // private _onRemoveSuggestion = (item: IPersonaProps): void => {
    //     const { peopleList, mostRecentlyUsed: mruState } = this.state;
    //     const indexPeopleList: number = peopleList.indexOf(item);
    //     const indexMostRecentlyUsed: number = mruState.indexOf(item);

    //     if (indexPeopleList >= 0) {
    //         const newPeople: IPersonaProps[] = peopleList
    //             .slice(0, indexPeopleList)
    //             .concat(peopleList.slice(indexPeopleList + 1));
    //         this.setState({ peopleList: newPeople });
    //     }

    //     if (indexMostRecentlyUsed >= 0) {
    //         const newSuggestedPeople: IPersonaProps[] = mruState
    //             .slice(0, indexMostRecentlyUsed)
    //             .concat(mruState.slice(indexMostRecentlyUsed + 1));
    //         this.setState({ mostRecentlyUsed: newSuggestedPeople });
    //     }
    // };


}
