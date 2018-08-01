import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { find, filter, sortBy, uniqBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IAddProjectProps } from './IAddProjectProps';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import "bootstrap/dist/css/bootstrap.min.css";
import { Button, Modal } from 'react-bootstrap';

import ProjectListTable from '../ProjectList/ProjectListTable';


import { Project } from "../ProjectList/ProjectList";


//Start: People Picker

import { BaseComponent, assign } from 'office-ui-fabric-react/lib/Utilities';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { IPersonaProps, Persona } from 'office-ui-fabric-react/lib/Persona';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    IBasePicker,
    ListPeoplePicker,
    NormalPeoplePicker,
    ValidationState
} from 'office-ui-fabric-react/lib/Pickers';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.types';
//import { people, mru } from './PeoplePickerExampleData';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Promise } from 'es6-promise';


//for react select
import Select from 'react-select';
import CreatableSelect from 'react-select/lib/Creatable';

export const colourOptions = [
    { value: 'ocean', label: 'Ocean', color: '#00B8D9' },
    { value: 'blue', label: 'Blue', color: '#0052CC', disabled: true },
    { value: 'purple', label: 'Purple', color: '#5243AA' },
    { value: 'red', label: 'Red', color: '#FF5630' },
    { value: 'orange', label: 'Orange', color: '#FF8B00' },
    { value: 'yellow', label: 'Yellow', color: '#FFC400' },
    { value: 'green', label: 'Green', color: '#36B37E' },
    { value: 'forest', label: 'Forest', color: '#00875A' },
    { value: 'slate', label: 'Slate', color: '#253858' },
    { value: 'silver', label: 'Silver', color: '#666666' },
];

const suggestionProps: IBasePickerSuggestionsProps = {
    //suggestionsHeaderText: 'Suggested People',
    // mostRecentlyUsedHeaderText: 'Suggested Contacts',
    // noResultsFoundText: 'No results found',
    loadingText: 'Loading',
    showRemoveButtons: true,
    suggestionsAvailableAlertText: 'People Picker Suggestions available',
    //suggestionsContainerAriaLabel: 'Suggested contacts',

};
const limitedSearchAdditionalProps: IBasePickerSuggestionsProps = {
    searchForMoreText: 'Load all Results',
    resultsMaximumNumber: 10,
    searchingText: 'Searching...',
};
const limitedSearchSuggestionProps: IBasePickerSuggestionsProps = assign(limitedSearchAdditionalProps, suggestionProps);

//End: People Picker

export default class AddProject extends React.Component<IAddProjectProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass: {},
    cloneProjectChecked: boolean,
    showModal: boolean,
    projectList: any,
    peopleList: any[],
    mostRecentlyUsed: IPersonaProps[];
    currentSelectedItems?: IPersonaProps[];
    delayResults?: boolean,
    currentPicker?: number | string,
    isDataSaved: boolean,
    showStatusDate: boolean,
    selectedOption: null,
    inputValue: any,
    value: any,
    tagOptions: any
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
            projectList: Array<Project>(),
            peopleList: peopleList,
            mostRecentlyUsed: [],
            currentSelectedItems: [],
            delayResults: false,
            currentPicker: 1,
            isDataSaved: false,
            showStatusDate: true,
            selectedOption: null,
            inputValue: '',
            value: [],
            tagOptions: []
        };
        this._showModal = this._showModal.bind(this);
        this._closeModal = this._closeModal.bind(this);
    }
    componentDidMount() {
        this._getAllSiteUsers();
        this.getAllProject();
        this.getAllProjectTags();
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
    getAllProjectTags() {
        sp.web.lists.getByTitle("Project Tags").items
            .select("Projects/ID", "Tag").expand("Projects")
            .get()
            .then((response) => {
                console.log(' all Project tag -', response);
                if (response.length > 0) {
                    let tempArray = {};
                    let tempList = [];
                    response.forEach(element => {
                        tempArray = {
                            value: element.Tag, label: element.Tag, color: '#00B8D9'
                        }
                        tempList.push(tempArray);
                    });
                    this.setState({ tagOptions: uniqBy(tempList, 'label') })
                }
            });
    }

    handleChange2 = (newValue: any, actionMeta: any) => {
        let fields = this.state.fields;
        fields['tags'] = newValue;
        this.setState(fields);
        console.group('Value Changed');
        console.log(newValue);

        console.log(`action: ${actionMeta.action}`);
        console.groupEnd();
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
            if (isChecked) {
                this.getProjectByName(this.state.projectList ? this.state.projectList[0].Project : '');
            } else if (this.props.id === undefined) {
                this.clearProjectInfo();
            }
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
        }
        else if (field === 'project') {
            let fields = this.state.fields;
            fields[field] = e.target.value;
            this.setState({ fields });
            this.getProjectByName(e.target.value);
        } else if (field === 'status') {
            if (e.target.value === 'On Hold') {
                let fields = this.state.fields;
                fields[field] = e.target.value;
                this.setState({ fields, showStatusDate: true });
            } else {
                let fields = this.state.fields;
                fields[field] = e.target.value;
                this.setState({ fields, showStatusDate: false });
            }
        } else if (field === 'statusdate') {
            let fields = this.state.fields;
            fields[field] = e;
        } else if (field === 'ownername') {
            let fields = this.state.fields;
            let ownerArray = [];
            e.forEach(element => {
                ownerArray.push(element.key);
            });
            fields[field] = ownerArray;
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
        if (!this.state.currentSelectedItems || this.state.currentSelectedItems.length === 0) {
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
        if (!fields["tags"] || fields["tags"].length === 0) {
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
        if (!fields["statusdate"] && this.state.showStatusDate && this.props.id) {
            formIsValid = false;
            errors["statusdate"] = "Cannot be empty";
            errorClass["statusdate"] = "classError";
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
    clearProjectInfo() {
        let fields = this.state.fields;
        fields["project"] = '';
        fields["projectname"] = '';
        fields["priority"] = '';
        fields["duedate"] = '';
        fields["ownername"] = '';
        fields["projectdescription"] = '';
        fields["startdate"] = '';
        fields["departmentspecific"] = false;
        fields["requringproject"] = false;
        fields["occurance"] = '';
        fields["cloneschedule"] = false;
        fields["clonedocuments"] = false;
        fields["clonerequirements"] = false;
        fields["clonecalender"] = false;
        this.setState({ fields });
    }
    projectSubmit(e) {
        e.preventDefault();
        if (this.handleValidation()) {
            let obj: any = this.state.fields;
            let fields = this.state.fields;
            let tempState: any = this.state.currentSelectedItems;
            let ownerArray = [];
            tempState.forEach(element => {
                ownerArray.push(element.key);
            });
            fields['ownername'] = ownerArray;
            this.setState({ fields });
            if (this.props.id) {
                sp.web.lists.getByTitle("Project").items.getById(this.props.id).update({
                    StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
                    DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : '',
                    //Status0Id: 2,
                    AssignedToId: { results: obj.ownername },
                    Priority: obj.priority ? obj.priority : 'Low',
                    Clone_x0020_Project: obj.cloneproject ? obj.cloneproject : false,
                    Clone_x0020_Documents: obj.clonedocuments ? obj.clonedocuments : false,
                    Clone_x0020_Requirements: obj.clonerequirements ? obj.clonerequirements : false,
                    Clone_x0020_Schedule: obj.cloneschedule ? obj.cloneschedule : false,
                    Clone_x0020_Calender: obj.clonecalender ? obj.clonecalender : false,
                    Body: obj.projectdescription ? obj.projectdescription : '',
                    Occurance: obj.occurance ? obj.occurance : 'Daily',
                    Recurring_x0020_Project: obj.requringproject ? obj.requringproject : false,
                    ProTypeDeptSpecific: obj.departmentspecific ? obj.departmentspecific : false,
                    On_x0020_Hold_x0020_Status: obj.status ? obj.status : 'On Hold',
                    On_x0020_Hold_x0020_Date: obj.statusdate ? new Date(obj.statusdate).toDateString() : null,

                }).then(i => {
                    this.state.fields['tags'].forEach(element => {
                        this.addProjectTagByTagName(element.value, this.props.id);
                    });
                });
            } else {
                sp.web.lists.getByTitle("Project").items.add({
                    Project: obj.projectname ? obj.projectname : '',
                    StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
                    DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : '',
                    AssignedToId: { results: obj.ownername },
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
                    this.state.fields['tags'].forEach(element => {
                        this.addProjectTagByTagName(element.value, response.data.Id);
                    });
                    // this._closePanel();
                    // this._showModal();
                });
            }
        } else {
            console.log("Form has errors.")
        }
    }
    addProjectTagByTagName(tagName, projectId) {
        var available = false;
        var project = [];
        
        var filter = "Tag eq" + "'" + tagName + "'";
        sp.web.lists.getByTitle("Project Tags").items
            .select("Projects/ID", "Tag").expand("Projects")
            .filter(filter)
            .getAll()
            .then((response) => {
                if (response.length > 0) {
                    response.forEach(element => {
                        if (element.Projects.length == 0) {
                            this.updateProjectTag(tagName, projectId, project, filter)
                        } else {
                            for (var i = 0; i < element.Projects.length; i++) {
                                project.push(element.Projects[i].ID)
                                if (element.Projects[i].ID == projectId) {
                                    available = true;
                                }
                            }
                            if (!available) {
                                this.updateProjectTag(tagName, projectId, project, filter)
                            }
                        }
                    });
                    console.log('Project tag 1 -', response);
                } else {

                    this.addProjectTag(tagName, projectId)
                }
            });
    }
    addProjectTag(tagName, projectID) {
        sp.web.lists.getByTitle("Project Tags").items.add({
            ProjectsId: { results: [projectID] },
            Tag: tagName
        }).then((response) => {
            if (this.props.id) {
                this._closePanel();
                this.props.parentMethod();
                //this.props.parentReopen();
            } else {
                this._closePanel();
                this._showModal();
            }
            console.log('Project team members added -', response);
        });
    }

    updateProjectTag(TagName, projectId, project, filter) {
        project.push(projectId);
        sp.web.lists.getByTitle("Project Tags").items.top(1).filter(filter).getAll().then((items: any[]) => {
            // see if we got something
            if (items.length > 0) {
                sp.web.lists.getByTitle("Project Tags").items.getById(items[0].Id).update({
                    ProjectsId: { results: project }
                }).then(result => {
                    // here you will have updated the item
                    if (this.props.id) {
                        this._closePanel();
                        this.props.parentMethod();
                        //this.props.parentReopen();
                    } else {
                        this._closePanel();
                        this._showModal();
                    }
                    console.log(JSON.stringify(result));
                });
            }
        });
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
    public render(): React.ReactElement<IAddProjectProps> {
        const { selectedOption } = this.state;
        const { inputValue, value } = this.state;

        let formControl = 'form-control';
        let paddingInputStyle = 'padding-input-style';
        const statusContent = this.props.id ?
            <div className="col-lg-6">
                <div className="form-group">
                    <label>Status</label>
                    <select ref="status" className={formControl + " " + (this.state.errorClass["status"] ? this.state.errorClass["status"] : '')}
                        onChange={this.handleChange.bind(this, "status")} value={this.state.fields["status"]}>
                        <option>On Hold</option>
                        <option>Resume </option>
                    </select>
                </div>
            </div> : null;
        const statusDate = (this.props.id && this.state.showStatusDate) ?
            <div className="col-lg-6">
                <div className="form-group">
                    <label>On Hold Date</label>
                    <DatePicker
                        placeholder="Select On Hold date"
                        onSelectDate={this.handleChange.bind(this, "statusdate")}
                        value={this.state.fields["statusdate"]}
                    />
                    <span className="error">{this.state.errors["statusdate"]}</span>
                </div>
            </div> : null
        const selectProjectContent = this.state.cloneProjectChecked ?
            <div className="col-lg-12">
                <div className="form-group">
                    <label>Select Project</label>
                    <select className="form-control" ref="project" onChange={this.handleChange.bind(this, "project")} value={this.state.fields["project"]}>
                        {this.state.projectList.map((obj) =>
                            <option key={obj.Project} value={obj.Project}>{obj.Project}</option>
                        )}
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
                                                                        <span className="calendar-style">
                                                                            {/* <i className="fas fa-user icon-style"></i> */}
                                                                            {/* <input ref="ownername" type="text" className={paddingInputStyle + " " + formControl + " " + (this.state.errorClass["ownername"] ? this.state.errorClass["ownername"] : '')} placeholder="Enter owners name"
                                                                                onChange={this.handleChange.bind(this, "ownername")} value={this.state.fields["ownername"]}>
                                                                            </input> */}
                                                                            {/* <NormalPeoplePicker
                                                                                onResolveSuggestions={this._onFilterChanged}
                                                                                onEmptyInputFocus={this._returnMostRecentlyUsed}
                                                                                getTextFromItem={this._getTextFromItem}
                                                                                pickerSuggestionsProps={suggestionProps}
                                                                                className={'ms-PeoplePicker'}
                                                                                key={'normal'}
                                                                                onRemoveSuggestion={this._onRemoveSuggestion}
                                                                                onValidateInput={this._validateInput}
                                                                                removeButtonAriaLabel={'Remove'}
                                                                                inputProps={{
                                                                                    onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                                                                                    onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
                                                                                    'aria-label': 'People Picker'
                                                                                }}
                                                                                //componentRef={this._resolveRef('_picker')}
                                                                                onInputChange={this._onInputChange}
                                                                                onChange={this.handleChange.bind(this, "ownername")}
                                                                                resolveDelay={300}
                                                                                defaultSelectedItems={selectedPeopleList}
                                                                            /> */}
                                                                            {this._renderControlledPicker()}
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
                                                                        {/* <input ref="tags" type="text" className={formControl + " " + (this.state.errorClass["tags"] ? this.state.errorClass["tags"] : '')} placeholder="Enter Tags"
                                                                            onChange={this.handleChange.bind(this, "tags")} value={this.state.fields["tags"]}>
                                                                        </input> */}
                                                                        <CreatableSelect
                                                                            isMulti
                                                                            onChange={this.handleChange2}
                                                                            options={this.state.tagOptions}
                                                                            value={this.state.fields["tags"]}
                                                                        />
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
                                                                {statusContent}
                                                                {statusDate}
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
            .select("Project", "StartDate", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/EMail", "AssignedTo/ID", "Priority", "Clone_x0020_Project", "Clone_x0020_Calender", "Clone_x0020_Documents", "Clone_x0020_Requirements", "Clone_x0020_Schedule", "Body", "Occurance", "Recurring_x0020_Project", "ProTypeDeptSpecific", "On_x0020_Hold_x0020_Date", "On_x0020_Hold_x0020_Status")
            .expand("Status0", "AssignedTo")
            .filter(filterString)
            .getAll()
            .then((response) => {
                console.log('Project by name', response);
                this.state.fields["project"] = response ? response[0].Project : '';
                this.state.fields["projectname"] = response ? response[0].Project : '';
                this.state.fields["priority"] = response ? response[0].Priority : '';
                this.state.fields["duedate"] = response ? new Date(response[0].DueDate) : '';
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
                this.state.fields["status"] = response ? response[0].On_x0020_Hold_x0020_Status : '';

                const selectedPeopleList: IPersonaWithMenu[] = [];
                const selectedTarget: IPersonaWithMenu = {};
                let tempSelectedPersona = {};
                if (response[0].AssignedTo.length > 0) {
                    response[0].AssignedTo.forEach(element => {
                        tempSelectedPersona = {
                            key: element.ID,
                            text: element.Title
                        }
                        //assign(selectedTarget, tempSelectedPersona);
                        selectedPeopleList.push(tempSelectedPersona);
                    });
                }
                // this.setState({
                //     assignedTo: response[0].AssignedTo
                // });
                this.setState({ currentSelectedItems: selectedPeopleList })
                this.state.fields["ownername"] = selectedPeopleList;

                if (response[0].Clone_x0020_Project) {
                    this.setState({ cloneProjectChecked: true });
                } else {
                    this.setState({ cloneProjectChecked: false });
                }
                if (response[0].On_x0020_Hold_x0020_Status === 'On Hold') {
                    this.state.fields["statusdate"] = response ? new Date(response[0].On_x0020_Hold_x0020_Date) : '';
                    this.setState({ showStatusDate: true });
                } else if (response[0].On_x0020_Hold_x0020_Status === null) {
                    this.state.fields["status"] = 'On Hold';
                    this.setState({ showStatusDate: true });
                } else {
                    this.state.fields["statusdate"] = '';
                    this.setState({ showStatusDate: false });
                }
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });
    }
    private getProjectTagsByProjectName(id): void {
        let filterString = "Projects/ID eq " + id;
        //let filterString = "Project/ID eq 1";
        sp.web.lists.getByTitle("Project Tags").items
            .select("Projects/ID", "Tag").expand("Projects")
            .filter(filterString)
            .get()
            .then((response) => {
                console.log('Project tag 1 -', response);
                let fields = this.state.fields;
                if (response.length > 0) {
                    var tempTag = {};
                    var tempTagList = [];
                    response.forEach(element => {
                        tempTag = { "value": element.Tag, "label": element.Tag, "color": "#00B8D9" };
                        tempTagList.push(tempTag);
                    });
                }
                fields["tags"] = tempTagList;
                this.setState({ fields });
            });
    }
    getAllProject() {
        // get Project Documents list items for all projects
        sp.web.lists.getByTitle("Project").items
            .select("Project", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "Priority").expand("Status0", "AssignedTo")
            .getAll()
            .then((response) => {
                console.log('projects', response);
                this.setState({ projectList: response });
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });
    }
    private getProjectByName(name) {
        let filter = "Project eq '" + name + "'";
        sp.web.lists.getByTitle("Project").items
            .select("Project", "StartDate", "DueDate", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/ID", "Priority", "Clone_x0020_Project", "Clone_x0020_Calender", "Clone_x0020_Documents", "Clone_x0020_Requirements", "Clone_x0020_Schedule", "Body", "Occurance", "Recurring_x0020_Project", "ProTypeDeptSpecific")
            .expand("Status0", "AssignedTo")
            .filter(filter)
            .getAll()
            .then((response) => {
                console.log('getProjectDetails', response);
                let fields = this.state.fields;
                fields["project"] = response ? response[0].Project : '';
                fields["projectname"] = response ? response[0].Project : '';
                fields["priority"] = response ? response[0].Priority : '';
                fields["duedate"] = response ? new Date(response[0].DueDate) : '';
                fields["ownername"] = response ? response[0].AssignedTo[0].Title : '';
                fields["projectdescription"] = response ? response[0].Body : '';
                fields["startdate"] = response ? new Date(response[0].StartDate) : '';
                fields["departmentspecific"] = response ? response[0].ProTypeDeptSpecific : false;
                fields["requringproject"] = response ? response[0].Recurring_x0020_Project : false;
                fields["occurance"] = response ? response[0].Occurance : '';
                fields["cloneschedule"] = response ? response[0].Clone_x0020_Schedule : false;
                fields["clonedocuments"] = response ? response[0].Clone_x0020_Documents : false;
                fields["clonerequirements"] = response ? response[0].Clone_x0020_Requirements : false;
                fields["clonecalender"] = response ? response[0].Clone_x0020_Calender : false;
                this.setState({ fields });
            }).catch((e: Error) => {
                alert(`There was an error : ${e.message}`);
            });
    }



    /*Start: People Picker Methods */
    private _getAllSiteUsers = (): void => {
        var reactHandler = this;
        sp.web.siteUsers.get().then(function (data) {
            const peopleList: IPersonaWithMenu[] = [];
            data.forEach((persona) => {
                let profileUrl = "https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=" +
                    persona.Email +
                    "&UA=0&size=HR64x64&sc=1531997060853";
                const target: IPersonaWithMenu = {};
                let tempPersona = {
                    key: persona.Id,
                    text: persona.Title,

                    imageUrl: persona.Email === undefined || persona.Email === '' ? null : profileUrl
                }
                assign(target, tempPersona);
                peopleList.push(target);

            });

            const mru: IPersonaProps[] = peopleList.slice(0, 5);
            reactHandler.setState({
                peopleList: peopleList,
                //mostRecentlyUsed: mru
            });
            //console.log('People : ' + peopleList);
        });
    };

    private _getTextFromItem(persona: IPersonaProps): string {
        return persona.text as string;
    }

    private _renderControlledPicker() {
        const controlledItems = [];
        for (let i = 0; i < 5; i++) {
            const item = this.state.peopleList[i];
            if (this.state.currentSelectedItems!.indexOf(item) === -1) {
                controlledItems.push(this.state.peopleList[i]);
            }
        }
        return (
            <div>
                <NormalPeoplePicker
                    onResolveSuggestions={this._onFilterChanged}
                    getTextFromItem={this._getTextFromItem}
                    pickerSuggestionsProps={suggestionProps}
                    className={'ms-PeoplePicker'}
                    key={'controlled'}
                    selectedItems={this.state.currentSelectedItems}
                    onChange={this._onItemsChange}
                    inputProps={{
                        onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                        onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called')
                    }}
                    //componentRef={this._resolveRef('_picker')}
                    resolveDelay={300}
                />
                {/* <label> Click to Add a person </label>
                {controlledItems.map((item, index) => (
                    <div key={index}>
                        <DefaultButton
                            className="controlledPickerButton"
                            // tslint:disable-next-line:jsx-no-lambda
                            onClick={() => {
                                this.setState({
                                    currentSelectedItems: this.state.currentSelectedItems!.concat([item])
                                });
                            }}
                        >
                            <Persona {...item} />
                        </DefaultButton>
                    </div>
                ))} */}
            </div>
        );
    }

    private _onItemsChange = (items: any[]): void => {
        this.setState({
            currentSelectedItems: items
        });
    };

    private _onSetFocusButtonClicked = (): void => {
        if (this._picker) {
            this._picker.focusInput();
        }
    };

    private _renderFooterText = (): JSX.Element => {
        return <div>No additional results</div>;
    };

    private _onRemoveSuggestion = (item: IPersonaProps): void => {
        const { peopleList, mostRecentlyUsed: mruState } = this.state;
        const indexPeopleList: number = peopleList.indexOf(item);
        const indexMostRecentlyUsed: number = mruState.indexOf(item);

        if (indexPeopleList >= 0) {
            const newPeople: IPersonaProps[] = peopleList
                .slice(0, indexPeopleList)
                .concat(peopleList.slice(indexPeopleList + 1));
            this.setState({ peopleList: newPeople });
        }

        if (indexMostRecentlyUsed >= 0) {
            const newSuggestedPeople: IPersonaProps[] = mruState
                .slice(0, indexMostRecentlyUsed)
                .concat(mruState.slice(indexMostRecentlyUsed + 1));
            this.setState({ mostRecentlyUsed: newSuggestedPeople });
        }
    };

    private _onItemSelected = (item: IPersonaProps): Promise<IPersonaProps> => {
        const processedItem = item;//Object.assign({}, item);
        processedItem.text = `${item.text} (selected)`;
        return new Promise<IPersonaProps>((resolve, reject) => setTimeout(() => resolve(processedItem), 250));
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

    private _returnMostRecentlyUsed = (currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> => {
        let { mostRecentlyUsed } = this.state;
        mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
        return this._filterPromise(mostRecentlyUsed);
    };

    private _returnMostRecentlyUsedWithLimit = (
        currentPersonas: IPersonaProps[]
    ): IPersonaProps[] | Promise<IPersonaProps[]> => {
        let { mostRecentlyUsed } = this.state;
        mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
        mostRecentlyUsed = mostRecentlyUsed.splice(0, 3);
        return this._filterPromise(mostRecentlyUsed);
    };

    private _onFilterChangedWithLimit = (
        filterText: string,
        currentPersonas: IPersonaProps[]
    ): IPersonaProps[] | Promise<IPersonaProps[]> => {
        return this._onFilterChanged(filterText, currentPersonas, 3);
    };

    private _filterPromise(personasToReturn: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        if (this.state.delayResults) {
            return this._convertResultsToPromise(personasToReturn);
        } else {
            return personasToReturn;
        }
    }

    private _listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
        if (!personas || !personas.length || personas.length === 0) {
            return false;
        }
        return personas.filter(item => item.text === persona.text).length > 0;
    }

    private _filterPersonasByText(filterText: string): IPersonaProps[] {
        return this.state.peopleList.filter(item => this._doesTextStartWith(item.text as string, filterText));
    }

    private _doesTextStartWith(text: string, filterText: string): boolean {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    }

    private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
        return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
    }

    private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
        return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
    }

    private _toggleDelayResultsChange = (toggleState: boolean): void => {
        this.setState({ delayResults: toggleState });
    };

    private _dropDownSelected = (option: IDropdownOption): void => {
        this.setState({ currentPicker: option.key });
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

    /**
     * Takes in the picker input and modifies it in whichever way
     * the caller wants, i.e. parsing entries copied from Outlook (sample
     * input: "Aaron Reid <aaron>").
     *
     * @param input The text entered into the picker.
     */
    private _onInputChange(input: string): string {
        const outlookRegEx = /<.*>/g;
        const emailAddress = outlookRegEx.exec(input);

        if (emailAddress && emailAddress[0]) {
            return emailAddress[0].substring(1, emailAddress[0].length - 1);
        }

        return input;
    }
    /*End: People Picker Methods */

}
