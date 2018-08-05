import * as React from "react";
import { sp, ItemAddResult, Web } from "@pnp/sp";
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
import { Link, Redirect } from 'react-router-dom';

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
//code from ashwini



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
export interface IHBCOwner {
    OwnerId: string | number;
    OwnerName: string;
    LoginName: string;
}
export interface ICloneProjectData {
    ID: string | number;
    Department: string;
    Status0: {
        ID: string | number;
        Status: string;
        StatusColor: string;
    }
    PercentComplete: string;
    AssignedTo: {
        ID: string | number;
        Title: string;
    }
    StartDate: string;
    DueDate: string;
    Body: string;
    Priority: string;
    ProjectTag: string;
    ProTypeDeptSpecific: boolean;
    RecurringProject: boolean;
    Occurance: string;
    Parent: string;
    IsActive: boolean;
    OnHoldStatus: boolean;
    OnHoldDate: boolean;
    ScheduleList: string;
    // Requirements: string;
    // ProjectDocument: string;
    // ProjectCalender: string;
    CalendarList: string;
    DocumentList: string;
    RequirementsList: string;
}
// export interface IListIDs {
//     ProjectList: string;
//     ProjectStatusColor: string;
//     TaskStatusColor: string;
//     Departments: string;
// }
export interface IRoleAssignments {
    RoleId: string;
    RoleName: string;
}

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
    tagOptions: any,
    // added code from ashwini
    // ViewersGroupId: string,
    // ContributersGroupId: string,
    // OwnersGroupId: string,
    HBCOwner: IHBCOwner[],
    CloneProjectData: ICloneProjectData[],
    ProjectList: string,
    TaskStatusColor: string,
    roleAssignments: IRoleAssignments[],
    savedProjectID: any,
    statusList: any,
    departmentList: any,
    showDepartment: boolean,
    isLoading: boolean
}> {
    private _picker: IBasePicker<IPersonaProps>;
    // Added by Ashwini
    public ViewersGroupId: string | number;
    public ContributersGroupId: string | number;
    public OwnersGroupId: string | number;
    //for permission
    public TaskObj: any;
    public ScheduleObj: any;
    public TaskCommentObj: any;
    public TaskCommentHisObj: any;
    public ProjectCommentObj: any;
    public ProjectCommentHisObj: any;
    public DocumentObj: any;
    public TeamMemberObj: any;
    public RequirementObj: any;
    public ProjectInfoObj: any;
    public ProjectCalendarObj: any;
    public TaskList = "_Task_List";
    public ScheduleList = "_Schedule_List";
    public ProjectDocument = "_Project_Document";
    public Requirements = "_Requirements";
    public ProjectTeamMembers = "_Project_Team_Members";
    public ProjectInfo = "_Project_Information";
    public ProjectCal = "_Project_Calender";
    public ProjectComments = "_Project_Comments";
    public ProjectCommentsHistory = "_Project_Comments_History";
    public TaskComments = "_Task_Comments";
    public TaskCommentsHistory = "_Task_Comments_History";
    public HBCAdminGrpID: string | number;
    public DepartmentHeadGrpID: string | number;
    public CEO_COOGrpID: string | number;
    public ProjectOwnerGrpID: string | number;


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
            showStatusDate: false,
            selectedOption: null,
            inputValue: '',
            value: [],
            tagOptions: [],
            // added code from ashwini
            // ViewersGroupId: '',
            // ContributersGroupId: '',
            // OwnersGroupId: '',
            HBCOwner: [],
            CloneProjectData: [],
            ProjectList: '',
            TaskStatusColor: '',
            roleAssignments: [],
            savedProjectID: 0,
            statusList: [],
            departmentList: [],
            showDepartment: false,
            isLoading: false,
        };
        this._showModal = this._showModal.bind(this);
        this._closeModal = this._closeModal.bind(this);
        this.handleBlurOnProjectName = this.handleBlurOnProjectName.bind(this);

        this.state.fields['status'] = false;
        this.state.fields['project'] = '';
    }
    componentDidMount() {

        this.getGroupID();

        this._getAllSiteUsers();
        this.getAllProject();
        this.getStatusList();
        this.getDepartmentList();
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

        // Added by Ashwini
        this.GetRoleDefinations();
        this.GetUsersFromHBCOwnerGroup();
        this.getListIDs();
    }
    componentWillReceiveProps(nextProps) {

    }

    // get group IDs by group name
    private getGroupID() {
        let reactHandler = this;
        sp.web.siteGroups.getByName("HBC Admin").get().then(function (result) {
            reactHandler.HBCAdminGrpID = result.Id;
            sp.web.siteGroups.getByName("Department Head").get().then(function (result) {
                reactHandler.DepartmentHeadGrpID = result.Id;
                sp.web.siteGroups.getByName("CEO_COO").get().then(function (result) {
                    reactHandler.CEO_COOGrpID = result.Id;
                    sp.web.siteGroups.getByName("Project Owner").get().then(function (result) {
                        reactHandler.ProjectOwnerGrpID = result.Id;
                    });
                });
            });
        });
    }

    private getListIDs() {

        sp.web.lists.getByTitle('Project').get()
            .then(result => {
                this.setState({
                    ProjectList: "'{" + result.Id + "}'"
                });

                sp.web.lists.getByTitle('Task Status Color').get()
                    .then(result => {
                        this.setState({
                            TaskStatusColor: "'{" + result.Id + "}'"
                        });

                    }).catch(err => {
                        console.log("Error while getting ID of Task Status Color List.", err);
                    });

            }).catch(err => {
                console.log("Error while getting ID of Project List.", err);
            });

    }

    // Get All Role Definations  
    private GetRoleDefinations() {
        let reactHandler = this;

        sp.web.roleDefinitions.get().then(result => {
            for (let i = 0; i < result.length; i++) {
                let id = result[i].Id;
                let name = result[i].Name;

                reactHandler.setState(prevState => ({
                    roleAssignments: [...prevState.roleAssignments, { RoleId: id, RoleName: name }]
                }));
            }
        }).catch(function (err) {
            console.log("Error: " + err);
        });
    }

    // Get Users Form HBC Owners Group
    private GetUsersFromHBCOwnerGroup() {
        let reactHandler = this;
        sp.web.siteGroups.getByName('HBC Dev Site Owners').users.get().then(function (result) {
            for (var i = 0; i < result.length; i++) {
                let ownerName = result[i].Title;
                let ownerId = result[i].Id;
                let loginName = result[i].LoginName;

                reactHandler.setState(prevState => ({
                    HBCOwner: [...prevState.HBCOwner, { OwnerName: ownerName, OwnerId: ownerId, LoginName: loginName }]
                }));
            }
        }).catch(function (err) {
            console.log("Group not found: " + err);
        });
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
    handleBlurOnProjectName() {
        console.log(this.state.fields['projectname']);
        let errors = this.state.errors;
        let errorClass = this.state.errorClass;
        if (this.state.fields['projectname']) {
            let flag = false;
            this.state.projectList.forEach(element => {
                if (element.Project.toLowerCase() === this.state.fields['projectname'].toLowerCase()) {
                    flag = true;
                }
            });
            if (flag) {
                errors["projectname"] = "Project name is already exist.";
                errorClass["projectname"] = "classError";
                this.setState({ errors: errors, errorClass: errorClass });
            } else {
                errors["projectname"] = "";
                errorClass["projectname"] = "";
                this.setState({ errors: errors, errorClass: errorClass });
            }
        }
    }
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
                //this.getProjectByName(this.state.projectList ? this.state.projectList[0].Project : '');
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
            if (isChecked) {
                this.setState({ fields, showDepartment: true });
            } else {
                this.setState({ fields, showDepartment: false });
            }
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
            if (e) {
                let fields = this.state.fields;
                fields[field] = e;
                this.setState({ fields, showStatusDate: true });
            } else {
                let fields = this.state.fields;
                fields[field] = e;
                fields['statusdate'] = null;
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

        } else if (field === 'projectoutline') {
            let fields = this.state.fields;
            fields[field] = e.target.files[0];
            this.setState({ fields });
        } else {
            let fields = this.state.fields;
            fields[field] = e.target.value;
            this.setState({ fields });
        }
    }
    removeAttachment(i, event) {
        console.log('index1', i);
        var result = confirm("Want to delete?");
        if (result) {
            let fields = this.state.fields;
            fields['projectoutline'].splice(i, 1);
            this.setState(fields);
        }
    }
    handleValidation() {
        let fields = this.state.fields;
        let errors = {};
        let errorClass = {};
        let formIsValid = true;

        //Name
        if (this.state.errors["projectname"]) {
            formIsValid = false;
            errors["projectname"] = "Project name is already exist.";
            errorClass["projectname"] = "classError";
        }
        if (!fields["projectname"]) {
            formIsValid = false;
            errors["projectname"] = "Cannot be empty";
            errorClass["projectname"] = "classError";
        }
        // if (!this.state.currentSelectedItems || this.state.currentSelectedItems.length === 0) {
        //     formIsValid = false;
        //     errors["ownername"] = "Cannot be empty";
        //     errorClass["ownername"] = "classError";
        // }
        // if (!fields["projectdescription"]) {
        //     formIsValid = false;
        //     errors["projectdescription"] = "Cannot be empty";
        //     errorClass["projectdescription"] = "classError";
        // }
        // if (!fields["startdate"]) {
        //     formIsValid = false;
        //     errors["startdate"] = "Cannot be empty";
        //     errorClass["startdate"] = "classError";
        // }
        // if (!fields["duedate"]) {
        //     formIsValid = false;
        //     errors["duedate"] = "Cannot be empty";
        //     errorClass["duedate"] = "classError";
        // }
        // if (!fields["tags"] || fields["tags"].length === 0) {
        //     formIsValid = false;
        //     errors["tags"] = "Cannot be empty";
        //     errorClass["tags"] = "classError";
        // }
        if ((!fields["project"] || fields["project"] === '') && this.state.cloneProjectChecked) {
            formIsValid = false;
            errors["project"] = "Please select Project Name";
            errorClass["project"] = "classError";
        }
        if (fields["startdate"] && fields["duedate"]) {
            if (fields["duedate"] < fields["startdate"]) {
                formIsValid = false;
                errors["duedate"] = "Due Date should always be greater than Start Date";
                errorClass["duedate"] = "classError";
            }
        }
        // if (!fields["statusdate"] && this.state.showStatusDate && this.props.id) {
        if (!fields["statusdate"] && this.state.showStatusDate) {
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
        let errors = this.state.errors;
        let errorClass = this.state.errorClass;
        const selectedPeopleList: IPersonaWithMenu[] = [];
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
        fields["tags"] = [];

        errors["projectname"] = "";
        errors["ownername"] = "";
        errors["startdate"] = "";
        errors["projectdescription"] = "";
        errors["duedate"] = "";
        errors["tags"] = "";

        errorClass["projectname"] = "";
        errorClass["projectdescription"] = "";


        this.setState({ currentSelectedItems: selectedPeopleList, fields, errors, errorClass });
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
                this.setState({ isLoading: true });
                sp.web.lists.getByTitle("Project").items.getById(this.props.id).update({
                    StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : null,
                    DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : null,
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
                    On_x0020_Hold_x0020_Status: obj.status ? obj.status : false,
                    On_x0020_Hold_x0020_Date: obj.statusdate && this.state.showStatusDate ? new Date(obj.statusdate).toDateString() : null,
                    Status0Id: obj.projectstatus ? obj.projectstatus : 1,
                    Risks: obj.risk ? obj.risk : 'Low',
                    DepartmentId: obj.departmentname && this.state.showDepartment ? obj.departmentname : 1

                }).then(i => {
                    this.setState({ isLoading: false });
                    this._closePanel();
                    this.props.parentMethod();
                    if (this.state.fields["projectoutline"].length > 0) {
                        console.log('Saving project outline....................');
                        i.item.attachmentFiles.add(this.state.fields["projectoutline"].name, this.state.fields["projectoutline"]);
                    }
                    if (this.state.fields['tags']) {
                        this.state.fields['tags'].forEach(element => {
                            this.addProjectTagByTagName(element.value, this.props.id);
                        });
                    }
                });
            } else {
                // sp.web.lists.getByTitle("Project").items.add({
                //     Project: obj.projectname ? obj.projectname : '',
                //     StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
                //     DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : '',
                //     AssignedToId: { results: obj.ownername },
                //     Priority: obj.priority ? obj.priority : 'Low',
                //     Body: obj.projectdescription ? obj.projectdescription : '',
                //     ProTypeDeptSpecific: obj.departmentspecific ? obj.departmentspecific : false,
                //     Recurring_x0020_Project: obj.requringproject ? obj.requringproject : false,
                //     Occurance: obj.occurance ? obj.occurance : 'Daily',
                //     Clone_x0020_Project: obj.cloneproject ? obj.cloneproject : false,
                //     Clone_x0020_Documents: obj.clonedocuments ? obj.clonedocuments : false,
                //     Clone_x0020_Requirements: obj.clonerequirements ? obj.clonerequirements : false,
                //     Clone_x0020_Schedule: obj.cloneschedule ? obj.cloneschedule : false,
                //     Clone_x0020_Calender: obj.clonecalender ? obj.clonecalender : false,
                //     //DepartmentId: 2,
                //     //Status0Id: 2

                // }).then((response) => {
                //     console.log('Item adding-', response);
                //     this.setState({ isDataSaved: true });
                //     this.state.fields['tags'].forEach(element => {
                //         this.addProjectTagByTagName(element.value, response.data.Id);
                //     });
                //     // this._closePanel();
                //     // this._showModal();
                // });
                this.setState({ isLoading: true });
                this.CreateProjectGroup();
                // this._closePanel();
                // this._showModal();
            }
        } else {
            console.log("Form has errors.")
        }
    }
    // Added new Code from Ashwini

    private CreateProjectGroup() {
        // let ProName = this.state.value;
        let ProName = this.state.fields['projectname'];
        let OwnersGroup = ProName + " Owners";
        let ContributersGroup = ProName + " Contributers";
        let ViewersGroup = ProName + " Viewers";
        let reactHandler = this;
        // Viewers
        sp.web.siteGroups.add({
            Title: ViewersGroup
        }).then(function (result) {
            console.log(ViewersGroup, " created !");
            // reactHandler.setState({
            //     ViewersGroupId: result.data.Id
            // });
            reactHandler.ViewersGroupId = result.data.Id;
            // Contributers
            sp.web.siteGroups.add({
                Title: ContributersGroup
            }).then(function (result) {
                console.log(ContributersGroup, " created !");
                // reactHandler.setState({
                //     ContributersGroupId: result.data.Id
                // });
                reactHandler.ContributersGroupId = result.data.Id;
                // Owners
                sp.web.siteGroups.add({
                    Title: OwnersGroup
                }).then(function (result) {
                    console.log(OwnersGroup, " created !");
                    // reactHandler.setState({
                    //     OwnersGroupId: result.data.Id
                    // });
                    reactHandler.OwnersGroupId = result.data.Id;
                    //Add User to group
                    reactHandler.AddHBCUsersToGroup(OwnersGroup, ProName);
                    // Add Item to project list & other list creation
                    let projectName = ProName.split(' ').join('_');
                    //if (reactHandler.state.fields['cloneproject'] !== true) {
                    reactHandler.CreateNewProject(projectName);
                    // }
                    // else {
                    //     reactHandler.getProjectDetails(projectName);
                    // }
                }).catch(e => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating " + OwnersGroup + " group: " + e);
                });
            }).catch(e => {
                this.setState({ isLoading: false });
                console.log("Error while creating " + ContributersGroup + " group: " + e);
            });
        }).catch(e => {
            this.setState({ isLoading: false });
            console.log("Error while creating " + ViewersGroup + " group: " + e);
        });
    }

    private AddHBCUsersToGroup(groupName, ProName) {
        let reactHandler = this;
        // add HBCOwner
        for (var i = 0; i < this.state.HBCOwner.length; i++) {
            let loginName = this.state.HBCOwner[i].LoginName

            sp.web.siteGroups.getByName(groupName).users.add(loginName)
                .then(function (d) {
                    console.log("HBCOwner added");

                }).catch(e => {
                    this.setState({ isLoading: false });
                    console.log("error while adding HBCOwner " + loginName + " to owner group");
                });
        }
    }
    // private AddPermissionsToTaskList(ListName, ProName) {

    //     this.TaskObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.TaskObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.TaskObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.TaskObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }

    // private AddPermissionsToScheduleList(ListName, ProName) {

    //     this.ScheduleObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.ScheduleObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.ScheduleObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.ScheduleObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }


    // private AddPermissionsToTaskCommentList(ListName, ProName) {

    //     this.TaskCommentObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.TaskCommentObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.TaskCommentObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.TaskCommentObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }


    // private AddPermissionsToTaskCommentHisList(ListName, ProName) {

    //     this.TaskCommentHisObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.TaskCommentHisObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.TaskCommentHisObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.TaskCommentHisObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }

    // private AddPermissionsToProjectCommentList(ListName, ProName) {

    //     this.ProjectCommentObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.ProjectCommentObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.ProjectCommentObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.ProjectCommentObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }


    // private AddPermissionsToProjectCommentHisList(ListName, ProName) {
    //     let reactHandler = this;

    //     this.ProjectCommentHisObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.ProjectCommentHisObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.ProjectCommentHisObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.ProjectCommentHisObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                     reactHandler.cloneListItems(ProName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }


    // private AddPermissionsToDocumentList(ListName, ProName) {

    //     this.DocumentObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.DocumentObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.DocumentObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.DocumentObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }



    // private AddPermissionsToTeamMemberList(ListName, ProName) {

    //     this.TeamMemberObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.TeamMemberObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.TeamMemberObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.TeamMemberObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }

    // private AddPermissionsToRequirementList(ListName, ProName) {

    //     this.RequirementObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.RequirementObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.RequirementObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.RequirementObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }

    // private AddPermissionsToProjectInfoList(ListName, ProName) {

    //     this.ProjectInfoObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.ProjectInfoObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.ProjectInfoObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.ProjectInfoObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }

    // private AddPermissionsToProjectCalendarList(ListName, ProName) {

    //     this.ProjectCalendarObj.breakRoleInheritance().then(res => {
    //         console.log("breakRoleInheritance for - ", ListName);

    //         this.ProjectCalendarObj.roleAssignments.add(this.OwnersGroupId, 1073741829).then(res => {
    //             console.log(ProName, " Owners - permissions added for -", ListName);

    //             this.ProjectCalendarObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => {
    //                 console.log(ProName, " Contributers - permissions added", ListName);

    //                 this.ProjectCalendarObj.roleAssignments.add(this.ViewersGroupId, 1073741924).then(res => {
    //                     console.log(ProName, " View Only - permissions added", ListName);

    //                 }).catch(err => {
    //                     console.log("Error while creating ", ProName, " View Only ");
    //                 });// View Only

    //             }).catch(err => {
    //                 console.log("Error while creating ", ProName, " Contributers ");
    //             });// Contribute

    //         }).catch(err => {
    //             console.log("Error while creating ", ProName, " Owners ");
    //         });; // Owners

    //     }).catch(err => {
    //         console.log("Error while breakRoleInheritance for list - ", ListName)
    //     });
    // }

    private getProjectDetails(ProName) {
        // let filter = "ID eq 175";
        let filter = "Project eq " + "'" + this.state.fields['project'] + "'";

        let ScheduleList;
        let DocumentList;
        let RequirementsList;
        let CalendarList;
        let reactHandler = this;

        sp.web.lists.getByTitle("Project").items
            .filter(filter)
            //.select("ID", "Project", "Schedule_x0020_List", "Requirements", "Project_x0020_Document", "Project_x0020_Calender")
            //  "Department",
            .select("ID", "Status0/ID", "Status0/Status", "Status0/Status_x0020_Color", "PercentComplete", "AssignedTo/ID", "AssignedTo/Title",
            "StartDate", "DueDate", "Body", "Priority", "ProTypeDeptSpecific", "Recurring_x0020_Project", "Occurance", "Parent",
            "IsActive", "On_x0020_Hold_x0020_Status", "On_x0020_Hold_x0020_Date", "Schedule_x0020_List", "Requirements", "Project_x0020_Document",
            "Project_x0020_Calender")
            .expand("Status0", "AssignedTo")
            .filter(filter)
            .getAll()
            .then((response) => {
                //console.log('getProjectDetails', response[0].Schedule_x0020_List);

                let ID = response[0].ID;
                let Department = response[0].Department;
                let Status0 = response[0].Status0;
                let PercentComplete = response[0].PercentComplete;
                let AssignedTo = response[0].AssignedTo;
                let StartDate = response[0].StartDate;
                let DueDate = response[0].DueDate;
                let Body = response[0].Body;
                let Priority = response[0].Priority;
                let ProjectTag = response[0].Project_x0020_Tag;
                let ProTypeDeptSpecific = response[0].ProTypeDeptSpecific;
                let RecurringProject = response[0].Recurring_x0020_Project;
                let Occurance = response[0].Occurance;
                let Parent = response[0].Parent;
                let IsActive = response[0].IsActive;
                let OnHoldStatus = response[0].On_x0020_Hold_x0020_Status;
                let OnHoldDate = response[0].On_x0020_Hold_x0020_Date;
                let ScheduleList = response[0].Schedule_x0020_List;
                let DocumentList = response[0].Project_x0020_Document;
                let RequirementsList = response[0].Requirements;
                let CalendarList = response[0].Project_x0020_Calender;

                reactHandler.setState(prevState => ({
                    CloneProjectData: [...prevState.CloneProjectData, {
                        ID: ID, Department: Department, Status0: Status0, PercentComplete: PercentComplete,
                        AssignedTo: AssignedTo, StartDate: StartDate, DueDate: DueDate, Body: Body, Priority: Priority,
                        ProjectTag: ProjectTag, ProTypeDeptSpecific: ProTypeDeptSpecific, RecurringProject: RecurringProject,
                        Occurance: Occurance, Parent: Parent, IsActive: IsActive, OnHoldStatus: OnHoldStatus, OnHoldDate: OnHoldDate,
                        ScheduleList: ScheduleList, DocumentList: DocumentList, RequirementsList: RequirementsList, CalendarList: CalendarList
                    }]
                }));

                // reactHandler.setState(prevState => ({
                //     CloneProjectData: [...prevState.CloneProjectData, {
                //         ScheduleList: ScheduleList, DocumentList: DocumentList, RequirementsList: RequirementsList, CalendarList: CalendarList
                //     }]
                // }));



                reactHandler.CreateCloneProject(ProName);

            }).catch(err => {
                console.log("Error while fetching project list items", err);
            });
    }

    private CreateCloneProject(ProName) {

        console.log("CloneProjectData", this.state.CloneProjectData);
        let obj: any = this.state.fields;
        let TaskListClmn = ProName + this.TaskList;
        let ScheduleListClmn = ProName + this.ScheduleList;
        let ProjectDocumentClmn = ProName + this.ProjectDocument;
        let RequirementsClmn = ProName + this.Requirements;
        let ProjectTeamMembersClmn = ProName + this.ProjectTeamMembers;
        let ProjectInfoClmn = ProName + this.ProjectInfo;
        let ProjectCalClmn = ProName + this.ProjectCal;
        let ProjectCommentsClmn = ProName + this.ProjectComments;
        let ProjectCommentsHistoryClmn = ProName + this.ProjectCommentsHistory;
        let TaskCommentsClmn = ProName + this.TaskComments;
        let TaskCommentsHistoryClmn = ProName + this.TaskCommentsHistory;

        // add an item to the list  
        sp.web.lists.getByTitle("Project").items.add({
            Title: "No Title",
            Project: ProName.split('_').join(' '),
            StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : null,
            DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : null,
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

            Task_x0020_List: TaskListClmn,
            Schedule_x0020_List: ScheduleListClmn,
            Project_x0020_Document: ProjectDocumentClmn,
            Requirements: RequirementsClmn,
            //Project_x0020_Tags: ProjectTags,
            Project_x0020_Team_x0020_Members: ProjectTeamMembersClmn,
            Project_x0020_Infromation: ProjectInfoClmn,
            Project_x0020_Calender: ProjectCalClmn,
            Project_x0020_Comments: ProjectCommentsClmn,
            Project_x0020_Comments_x0020_His: ProjectCommentsHistoryClmn,
            Task_x0020_Comments: TaskCommentsClmn,
            Task_x0020_Comments_x0020_Histor: TaskCommentsHistoryClmn
        }).then((iar: ItemAddResult) => {
            this.CreateTaskList(ProName);
        }).catch(err => {
            console.log("Error while cloning project", ProName, " - ", err);
        });
    }

    private CreateNewProject(ProName) {

        let obj: any = this.state.fields;
        let TaskListClmn = ProName + this.TaskList;
        let ScheduleListClmn = ProName + this.ScheduleList;
        let ProjectDocumentClmn = ProName + this.ProjectDocument;
        let RequirementsClmn = ProName + this.Requirements;
        let ProjectTeamMembersClmn = ProName + this.ProjectTeamMembers;
        let ProjectInfoClmn = ProName + this.ProjectInfo;
        let ProjectCalClmn = ProName + this.ProjectCal;
        let ProjectCommentsClmn = ProName + this.ProjectComments;
        let ProjectCommentsHistoryClmn = ProName + this.ProjectCommentsHistory;
        let TaskCommentsClmn = ProName + this.TaskComments;
        let TaskCommentsHistoryClmn = ProName + this.TaskCommentsHistory;

        // add an item to the list  
        sp.web.lists.getByTitle("Project").items.add({
            Project: ProName.split('_').join(' '),
            StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : null,
            DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : null,
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
            Status0Id: obj.projectstatus ? obj.projectstatus : 1,
            Risks: obj.risk ? obj.risk : 'Low',
            On_x0020_Hold_x0020_Status: obj.status ? obj.status : false,
            On_x0020_Hold_x0020_Date: obj.statusdate && this.state.showStatusDate ? new Date(obj.statusdate).toDateString() : null,
            DepartmentId: obj.departmentname && this.state.showDepartment ? obj.departmentname : 1,


            Title: "No Title",

            Task_x0020_List: TaskListClmn,
            Schedule_x0020_List: ScheduleListClmn,
            Project_x0020_Document: ProjectDocumentClmn,
            Requirements: RequirementsClmn,
            //Project_x0020_Tags: ProjectTags,
            Project_x0020_Team_x0020_Members: ProjectTeamMembersClmn,
            Project_x0020_Infromation: ProjectInfoClmn,
            Project_x0020_Calender: ProjectCalClmn,
            Project_x0020_Comments: ProjectCommentsClmn,
            Project_x0020_Comments_x0020_His: ProjectCommentsHistoryClmn,
            Task_x0020_Comments: TaskCommentsClmn,
            Task_x0020_Comments_x0020_Histor: TaskCommentsHistoryClmn
        }).then((iar: ItemAddResult) => {

            if (this.state.fields["projectoutline"]) {
                console.log('Saving project outline....................');
                iar.item.attachmentFiles.add(this.state.fields["projectoutline"].name, this.state.fields["projectoutline"]);
            }
            this.CreateTaskList(ProName);
            this.setState({ isDataSaved: true, savedProjectID: iar.data.Id });
            if (this.state.fields['tags']) {
                this.state.fields['tags'].forEach(element => {
                    this.addProjectTagByTagName(element.value, iar.data.Id);
                });
            }
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while adding items for ", ProName, " to Project List -", err);
        });
    }

    private CreateTaskList(ProName) {
        //let spWeb = new Web(this.context.pageContext.site.absoluteUrl);
        let spEnableCT = false;
        let reactHandler = this;
        let TaskList = ProName + this.TaskList;
        let TaskListDesc = TaskList + " Description";
        let TaskTemplateId = 171;


        sp.web.lists.add(TaskList, TaskListDesc, TaskTemplateId, spEnableCT).then(function (splist) {
            console.log(TaskList, " created successfuly !");

            reactHandler.AddTaskListColumns(TaskList, ProName, sp.web, spEnableCT, splist.list);


        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating task List, Error -", err);
        });
    }


    private AddTaskListColumns(ListName, ProName, spWeb, spEnableCT, list) {


        let ScheduleList = ProName + this.ScheduleList;
        this.TaskObj = list;

        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + `/>`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let Status = `<Field Name="Status" DisplayName="Status" Type="Lookup" Required="FALSE" ShowField="Status" List=` + this.state.TaskStatusColor + ` />`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Status).then(res => {
            console.log("Status created in list ", ListName);

            let Duration = "<Field Name='Duration' StaticName='Duration' DisplayName='Duration' Type='Text'  />";
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Duration).then(res => {
                console.log("Duration created in list ", ListName);

                let Comment = `<Field Name='Comment' StaticName='Comment' DisplayName='Comment' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Comment).then(res => {
                    console.log("Comment created in list ", ListName);
                    this.AddPermissionsToTaskList(ListName, ProName);
                    this.CreateScheduleList(ScheduleList, ProName, spWeb, spEnableCT);

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column Comment - ", " in list -", ListName, " Error -", err);
                }); //Comment

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Duration - ", " in list -", ListName, " Error -", err);
            }); //Duration

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Status - ", " in list -", ListName, " Error -", err);
        }); //Status
        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project
    }


    private CreateScheduleList(ScheduleList, ProName, spWeb, spEnableCT) {
        let reactHandler = this;
        let ScheduleListDesc = ScheduleList + " Description";
        let TaskTemplateId = 171;


        spWeb.lists.add(ScheduleList, ScheduleListDesc, TaskTemplateId, spEnableCT).then(function (splist) {
            console.log(ScheduleList, " created successfuly !");
            let ScheduleListID = "'{" + splist.data.Id + "}'";
            reactHandler.AddScheduleListIDColumns(ScheduleList, ProName, spWeb, spEnableCT, ScheduleListID, splist.list);

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Schedule List, Error -", err);
        });
    }

    private AddScheduleListIDColumns(ListName, ProName, spWeb, spEnableCT, ScheduleListID, list) {


        let TaskComments = ProName + this.TaskComments;
        this.ScheduleObj = list;


        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + `/>`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let Status = `<Field Name="Status" DisplayName="Status" Type="Lookup" Required="FALSE" ShowField="Status" List=` + this.state.TaskStatusColor + ` />`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Status).then(res => {
            console.log("Status created in list ", ListName);

            let Duration = "<Field Name='Duration' StaticName='Duration' DisplayName='Duration' Type='Text'  />";
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Duration).then(res => {
                console.log("Duration created in list ", ListName);

                let Comment = `<Field Name='Comment' StaticName='Comment' DisplayName='Comment' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Comment).then(res => {
                    console.log("Comment created in list ", ListName);
                    this.AddPermissionsToScheduleList(ListName, ProName);
                    this.CreateTaskComments(TaskComments, ProName, spWeb, spEnableCT, ScheduleListID);

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column Comment - ", " in list -", ListName, " Error -", err);
                }); //Comment

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Duration - ", " in list -", ListName, " Error -", err);
            }); //Duration

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Status - ", " in list -", ListName, " Error -", err);
        }); //Status
        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project
    }

    private CreateTaskComments(TaskComments, ProName, spWeb, spEnableCT, ScheduleListID) {
        var reactHandler = this;
        let TaskCommentsDesc = TaskComments + " Description";
        let TaskCommentsTemplateId = 100;
        // Create Project Calender List & Columns
        spWeb.lists.add(TaskComments, TaskCommentsDesc, TaskCommentsTemplateId, spEnableCT).then(function (splist) {
            console.log(TaskComments, " created successfuly !");
            let TaskCommentsID = "'{" + splist.data.Id + "}'";
            reactHandler.AddTaskCommentsColumns(TaskComments, ProName, spWeb, spEnableCT, ScheduleListID, TaskCommentsID, splist.list);

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Task Comments List, Error -", err);
        });
    }


    private AddTaskCommentsColumns(ListName, ProName, spWeb, spEnableCT, ScheduleListID, TaskCommentsID, list) {


        let TaskCommentsHistory = ProName + this.TaskCommentsHistory;
        this.TaskCommentObj = list;

        let TaskName = `<Field Name="Task Name" DisplayName="Task Name" Type="Lookup" Required="FALSE" ShowField="Title" List=` + ScheduleListID + ` />`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(TaskName).then(res => {
            console.log("Task created in list ", ListName);

            let Comment = `<Field Name='Comment' StaticName='Comment' DisplayName='Comment' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Comment).then(res => {
                console.log("Comment created in list ", ListName);

                this.AddPermissionsToTaskCommentList(ListName, ProName);
                this.CreateTaskCommentHistory(TaskCommentsHistory, ProName, spWeb, spEnableCT, TaskCommentsID);
            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Comment - ", " in list -", ListName, " Error -", err);
            }); //Comment

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column task - ", " in list -", ListName, " Error -", err);
        }); //Task Name
    }

    private CreateTaskCommentHistory(TaskCommentsHistory, ProName, spWeb, spEnableCT, TaskCommentsID) {
        var reactHandler = this;
        let TaskCommentsHistoryDesc = TaskCommentsHistory + " Description";
        let TaskCommentsHistoryTemplateId = 100;
        // Create Project Calender List & Columns
        spWeb.lists.add(TaskCommentsHistory, TaskCommentsHistoryDesc, TaskCommentsHistoryTemplateId, spEnableCT, ).then(function (splist) {
            console.log(TaskCommentsHistory, " created successfuly !");
            reactHandler.AddTaskCommentsHistoryColumns(TaskCommentsHistory, ProName, spWeb, spEnableCT, TaskCommentsID, splist.list);

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Task Comments History List, Error -", err);
        });
    }

    private AddTaskCommentsHistoryColumns(ListName, ProName, spWeb, spEnableCT, TaskCommentsID, list) {


        let ProjectDocument = ProName + this.ProjectDocument;
        this.TaskCommentHisObj = list;

        let TaskCommentID = `<Field Name="Task Comment ID" DisplayName="Task Comment ID" Type="Lookup" Required="FALSE" ShowField="ID" List=` + TaskCommentsID + `/>`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(TaskCommentID).then(res => {
            console.log("Task Comment ID created in list ", ListName);

            let Comment = `<Field Name='Comment' StaticName='Comment' DisplayName='Comment' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Comment).then(res => {
                console.log("Comment created in list ", ListName);

                let IsDeleted = `<Field Name='IsDeleted' StaticName='IsDeleted' DisplayName='IsDeleted' Type='Boolean'><Default>0</Default></Field>`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(IsDeleted).then(res => {
                    console.log("IsDeleted created in list ", ListName);

                    this.AddPermissionsToTaskCommentHisList(ListName, ProName);
                    this.CreateProjectDocument(ProName, ProjectDocument, spWeb, spEnableCT);
                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column IsDeleted - ", " in list -", ListName, " Error -", err);
                }); //IsDeleted 

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Comment - ", " in list -", ListName, " Error -", err);
            }); //Comment

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Task Comment ID - ", " in list -", ListName, " Error -", err);
        }); //Task Comment ID
    }

    private CreateProjectDocument(ProName, ProjectDocument, spWeb, spEnableCT) {
        var reactHandler = this;
        let ProjectDocumentDesc = ProjectDocument + " Description";
        let DocTemplateId = 101;

        spWeb.lists.add(ProjectDocument, ProjectDocumentDesc, DocTemplateId, spEnableCT).then(function (splist) {
            console.log(ProjectDocument, " created successfuly !");
            reactHandler.AddProjectDocColumns(ProjectDocument, ProName, spWeb, spEnableCT, splist.list);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Project Doc List, Error -", err);
        });
    }

    private AddProjectDocColumns(ListName, ProName, spWeb, spEnableCT, list) {
        //let Risks = ProName + " Risks";

        let Requirements = ProName + this.Requirements;

        this.DocumentObj = list;

        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + ` />`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let Owner = `<Field Name='Owner' StaticName='Owner' DisplayName='Owner' Type='User'/>`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Owner).then(res => {
            console.log("Owner created in list ", ListName);

            this.AddPermissionsToDocumentList(ListName, ProName);
            this.CreateRequirements(ProName, Requirements, spWeb, spEnableCT);

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Owner - ", " in list -", ListName, " Error -", err);
        }); //Owner

        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project
    }

    private CreateRequirements(ProName, Requirements, spWeb, spEnableCT) {
        var reactHandler = this;
        let RequirementsDesc = Requirements + " Description";
        let RequirementsTemplateId = 100;

        spWeb.lists.add(Requirements, RequirementsDesc, RequirementsTemplateId, spEnableCT).then(function (splist) {
            console.log(Requirements, " created successfuly !");
            // reactHandler.AddRequirementColumns(Requirements);
            reactHandler.AddRequirementColumns(Requirements, ProName, spWeb, spEnableCT, splist.list);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Requirement List, Error -", err);
        });
    }

    // private AddRequirementColumns(ListName){
    private AddRequirementColumns(ListName, ProName, spWeb, spEnableCT, list) {


        let ProjectTeamMembers = ProName + this.ProjectTeamMembers;

        this.RequirementObj = list;

        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + ` />`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let Requirement = `<Field Name='Requirement' StaticName='Requirement' DisplayName='Requirement' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Requirement).then(res => {
            console.log("Requirement created in list ", ListName);

            let Approver = `<Field Name='Approver' StaticName='Approver' DisplayName='Approver' Type='User'/>`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Approver).then(res => {
                console.log("Approver created in list ", ListName);

                let ApporvalStatus = `<Field DisplayName='Apporval Status' Name='ApporvalStatus' StaticName='Apporval Status' Type='Choice' Format='Dropdown'>
                                                            <Default>Pending</Default>
                                                            <CHOICES>
                                                               <CHOICE>Pending</CHOICE>
                                                               <CHOICE>Approved</CHOICE>
                                                               <CHOICE>Rejected</CHOICE>
                                                            </CHOICES>
                                                        </Field>`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(ApporvalStatus).then(res => {
                    console.log("Apporval Status created in list ", ListName);

                    let Efforts = `<Field Name='Efforts' StaticName='Efforts' DisplayName='Efforts' Type='Number'/>`;
                    sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Efforts).then(res => {
                        console.log("Efforts created in list ", ListName);

                        let Resources = `<Field Name='Resources' StaticName='Resources' DisplayName='Resources' Type='Number'/>`;
                        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Resources).then(res => {
                            console.log("Resources created in list ", ListName);

                            let ImpactOnTimelines = `<Field Name='ImpactonTimelines' StaticName='ImpactonTimelines' DisplayName='Impact On Timelines' Type='Boolean'><Default>0</Default></Field>`;
                            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(ImpactOnTimelines).then(res => {
                                console.log("ImpactOnTimelines created in list ", ListName);


                                this.AddPermissionsToRequirementList(ListName, ProName);
                                this.CreateProjectTeamMembers(ProName, ProjectTeamMembers, spWeb, spEnableCT);
                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while creating column ImpactOnTimelines - ", " in list -", ListName, " Error -", err);
                            }); //ImpactOnTimelines  

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while creating column Resources - ", " in list -", ListName, " Error -", err);
                        }); //Resources     

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while creating column  Efforts - ", " in list -", ListName, " Error -", err);
                    }); //Efforts

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column Apporval Status - ", " in list -", ListName, " Error -", err);
                }); //Apporval Status

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Approver - ", " in list -", ListName, " Error -", err);
            }); //Approver

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Requirement - ", " in list -", ListName, " Error -", err);
        }); //Requirement
        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project

    }

    private CreateProjectTeamMembers(ProName, ProjectTeamMembers, spWeb, spEnableCT) {
        var reactHandler = this;
        let TeamMembersDesc = ProjectTeamMembers + " Description";
        let TeamMembersTemplateId = 100;

        spWeb.lists.add(ProjectTeamMembers, TeamMembersDesc, TeamMembersTemplateId, spEnableCT).then(function (splist) {
            console.log(ProjectTeamMembers, " created successfuly !");
            reactHandler.AddTeamMemberColumns(ProjectTeamMembers, ProName, spWeb, spEnableCT, splist.list);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Team Members List, Error -", err);
        });
    }

    // private AddTeamMemberColumns(ListName){
    private AddTeamMemberColumns(ListName, ProName, spWeb, spEnableCT, list) {

        let ProjectInfo = ProName + this.ProjectInfo;

        this.TeamMemberObj = list;

        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + ` />`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let TeamMember = `<Field Name='TeamMember' StaticName='TeamMember' DisplayName='Team Member' Type='User'/>`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(TeamMember).then(res => {
            console.log("TeamMember created in list ", ListName);

            let StartDate = `<Field Name='StartDate' StaticName='StartDate' DisplayName='Start Date' Type='DateTime' Format='DateOnly' ></Field>`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(StartDate).then(res => {
                console.log("StartDate created in list ", ListName);

                let EndDate = `<Field Name='EndDate' StaticName='EndDate' DisplayName='End Date' Type='DateTime' Format='DateOnly' ></Field>`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(EndDate).then(res => {
                    console.log("EndDate created in list ", ListName);

                    let Status = `<Field DisplayName='Status' Name='Status' StaticName='Status' Type='Choice' Format='Dropdown'>
                                                            <Default>Active</Default>
                                                            <CHOICES>
                                                               <CHOICE>Active</CHOICE>
                                                               <CHOICE>Inactive</CHOICE>
                                                            </CHOICES>
                                                        </Field>`;
                    sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Status).then(res => {
                        console.log("Status created in list ", ListName);

                        this.AddPermissionsToTeamMemberList(ListName, ProName);
                        this.CreateProjectInfo(ProName, ProjectInfo, spWeb, spEnableCT);

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while creating column Status - ", " in list -", ListName, " Error -", err);
                    }); //Status

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column EndDate - ", " in list -", ListName, " Error -", err);
                }); //EndDate

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column StartDate - ", " in list -", ListName, " Error -", err);
            }); //StartDate

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column TeamMember - ", " in list -", ListName, " Error -", err);
        }); //TeamMember

        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project

    }

    private CreateProjectInfo(ProName, ProjectInfo, spWeb, spEnableCT) {
        var reactHandler = this;
        let ProjectInfoDesc = ProjectInfo + " Description";
        let ProjectInfoTemplateId = 100;

        spWeb.lists.add(ProjectInfo, ProjectInfoDesc, ProjectInfoTemplateId, spEnableCT).then(function (splist) {
            console.log(ProjectInfo, " created successfuly !");
            reactHandler.AddProjectInfoColumns(ProjectInfo, ProName, spWeb, spEnableCT, splist.list);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Team Members List, Error -", err);
        });
    }

    // private AddProjectInfoColumns(ListName){
    private AddProjectInfoColumns(ListName, ProName, spWeb, spEnableCT, list) {

        let ProjectCal = ProName + this.ProjectCal;

        this.ProjectInfoObj = list;

        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + ` />`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let Owner = `<Field Name='Owner' StaticName='Owner' DisplayName='Owner' Type='User'/>`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Owner).then(res => {
            console.log("Owner created in list ", ListName);

            let Roles_Responsibility = `<Field Name='Roles_Responsibility' StaticName='Roles_Responsibility' DisplayName='Roles_Responsibility' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Roles_Responsibility).then(res => {
                console.log("Roles_Responsibility created in list ", ListName);

                this.AddPermissionsToProjectInfoList(ListName, ProName);
                this.CreateProjectCalender(ProName, ProjectCal, spWeb, spEnableCT);

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Roles_Responsibility - ", " in list -", ListName, " Error -", err);
            }); //Roles_Responsibility

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Owner - ", " in list -", ListName, " Error -", err);
        }); //Owner

        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project
    }

    private CreateProjectCalender(ProName, ProjectCalender, spWeb, spEnableCT) {
        var reactHandler = this;
        let ProjectCalenderDesc = ProjectCalender + " Description";
        let ProjectCalenderTemplateId = 106;

        let ProjectComments = ProName + this.ProjectComments;

        spWeb.lists.add(ProjectCalender, ProjectCalenderDesc, ProjectCalenderTemplateId, spEnableCT).then(function (splist) {
            console.log(ProjectCalender, " created successfuly !");
            //reactHandler.AddProjectCalColumns(ProjectCalender, ProName, spWeb, spEnableCT);
            reactHandler.ProjectCalendarObj = splist.list;
            reactHandler.AddPermissionsToProjectCalendarList(ProjectCalender, ProName);
            reactHandler.CreateProjectComments(ProjectComments, ProName, spWeb, spEnableCT);
            //reactHandler.AddPermissionsToList(ProjectCalender, ProName, splist.list);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Project Calender List, Error -", err);
        });
    }

    // // private AddProjectCalColumns(ListName){
    // private AddProjectCalColumns(ListName, ProName, spWeb, spEnableCT) {
    //     let ProjectComments = ProName + " Project Comments";

    //     let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + ` />`;
    //     sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
    //         console.log("Project created in list ", ListName);
    //         this.CreateProjectComments(ProjectComments, ProName, spWeb, spEnableCT);
    //     }).catch(err => {
    //         console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
    //     }); //Project

    //     console.log("Operation Done!!!!");
    // }

    //Methods Added by ankit starts

    private CreateProjectComments(ProjectComments, ProName, spWeb, spEnableCT) {
        var reactHandler = this;
        let ProjectCommentDesc = ProjectComments + " Description";
        let ProjectCommentTemplateId = 100;
        // Create Project Comment List & Columns
        spWeb.lists.add(ProjectComments, ProjectCommentDesc, ProjectCommentTemplateId, spEnableCT).then(function (splist) {
            console.log(ProjectComments, " created successfuly !");
            let ProjectListID = "{" + splist.data.Id + "}";
            reactHandler.AddProjectCommentsColumns(ProjectComments, ProName, spWeb, spEnableCT, ProjectListID, splist.list);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Project Comments List, Error -", err);
        });
    }

    private AddProjectCommentsColumns(ListName, ProName, spWeb, spEnableCT, ProjectListID, list) {

        let ProjectCommentsHistory = ProName + this.ProjectCommentsHistory;

        this.ProjectCommentObj = list;

        // let Project = `<Field Name="Project" DisplayName="Project" Type="Lookup" Required="FALSE" ShowField="Project" List=` + this.state.ProjectList + ` />`;
        // sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Project).then(res => {
        //     console.log("Project created in list ", ListName);

        let Comment = `<Field Name='Comment' StaticName='Comment' DisplayName='Comment' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Comment).then(res => {
            console.log("Comment created in list ", ListName);

            let CommentID = `<Field Name='CommentID' StaticName='CommentID' DisplayName='CommentID' Type='Text' Sortable='FALSE' />`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(CommentID).then(res => {
                console.log("Comment ID created in list ", ListName);

                let IsDeleted = `<Field Name='IsDeleted' StaticName='IsDeleted' DisplayName='IsDeleted' Type='Boolean'><Default>0</Default></Field>`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(IsDeleted).then(res => {
                    console.log("IsDeleted created in list ", ListName);

                    this.AddPermissionsToProjectCommentList(ListName, ProName);
                    this.CreateProjectCommentHistory(ProjectCommentsHistory, ProName, spWeb, spEnableCT, ProjectListID);
                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column IsDeleted - ", " in list -", ListName, " Error -", err);
                }); //IsDeleted 
            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Comment ID - ", " in list -", ListName, " Error -", err);
            }); //CommentID
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column Comment - ", " in list -", ListName, " Error -", err);
        }); //Comment
        // }).catch(err => {
        //     console.log("Error while creating column Project - ", " in list -", ListName, " Error -", err);
        // }); //Project
    }

    private CreateProjectCommentHistory(ProjectCommentsHistory, ProName, spWeb, spEnableCT, ProjectListID) {
        var reactHandler = this;
        let ProjectCommentsHistoryDesc = ProjectCommentsHistory + " Description";
        let ProjectCommentsHistoryTemplateId = 100;
        // Create Project Comment History List & Columns
        spWeb.lists.add(ProjectCommentsHistory, ProjectCommentsHistoryDesc, ProjectCommentsHistoryTemplateId, spEnableCT).then(function (splist) {
            console.log(ProjectCommentsHistory, " created successfuly !");
            reactHandler.AddProjectCommentsHistoryColumns(ProjectCommentsHistory, ProName, spWeb, spEnableCT, ProjectListID, splist.list);

            //reactHandler.cloneListItems(ProName);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating Project Comments History List, Error -", err);
        });
    }

    private AddProjectCommentsHistoryColumns(ListName, ProName, spWeb, spEnableCT, ProjectListID, list) {


        this.ProjectCommentHisObj = list;

        let CommentID = `<Field Name="CommentID" DisplayName="CommentID" Type="Lookup" Required="FALSE" ShowField="CommentID" List="` + ProjectListID + `" />`;
        sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(CommentID).then(res => {
            console.log("CommentID created in list ", ListName);

            let Comment = `<Field Name='Comment' StaticName='Comment' DisplayName='Comment' Type='Note' NumLines='6' RichText='FALSE' Sortable='FALSE' />`;
            sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(Comment).then(res => {
                console.log("Comment created in list ", ListName);

                let IsDeleted = `<Field Name='IsDeleted' StaticName='IsDeleted' DisplayName='IsDeleted' Type='Boolean'><Default>0</Default></Field>`;
                sp.web.lists.getByTitle(ListName).fields.createFieldAsXml(IsDeleted).then(res => {
                    console.log("IsDeleted created in list ", ListName);

                    this.AddPermissionsToProjectCommentHisList(ListName, ProName);

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while creating column IsDeleted - ", " in list -", ListName, " Error -", err);
                }); //IsDeleted

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while creating column Comment - ", " in list -", ListName, " Error -", err);
            }); //Comment

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while creating column CommentID - ", " in list -", ListName, " Error -", err);
        }); //CommentID
    }

    // Add permissions to list
    private AddPermissionsToTaskList(ListName, ProName) {
        this.TaskObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.TaskObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.TaskObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.TaskObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.TaskObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.TaskObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }

    private AddPermissionsToScheduleList(ListName, ProName) {

        this.ScheduleObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.ScheduleObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.ScheduleObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.ScheduleObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.ScheduleObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.ScheduleObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }


    private AddPermissionsToTaskCommentList(ListName, ProName) {

        this.TaskCommentObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.TaskCommentObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.TaskCommentObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.TaskCommentObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.TaskCommentObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.TaskCommentObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }


    private AddPermissionsToTaskCommentHisList(ListName, ProName) {

        this.TaskCommentHisObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.TaskCommentHisObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.TaskCommentHisObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.TaskCommentHisObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.TaskCommentHisObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.TaskCommentHisObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }

    private AddPermissionsToProjectCommentList(ListName, ProName) {

        this.ProjectCommentObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.ProjectCommentObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.ProjectCommentObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.ProjectCommentObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.ProjectCommentObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.ProjectCommentObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }



    private AddPermissionsToDocumentList(ListName, ProName) {

        this.DocumentObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.DocumentObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.DocumentObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.DocumentObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.DocumentObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.DocumentObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }



    private AddPermissionsToTeamMemberList(ListName, ProName) {

        this.TeamMemberObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.TeamMemberObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.TeamMemberObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.TeamMemberObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.TeamMemberObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.TeamMemberObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }

    private AddPermissionsToRequirementList(ListName, ProName) {

        this.RequirementObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.RequirementObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.RequirementObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.RequirementObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.RequirementObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.RequirementObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp


        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }

    private AddPermissionsToProjectInfoList(ListName, ProName) {

        this.ProjectInfoObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.ProjectInfoObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.ProjectInfoObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.ProjectInfoObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.ProjectInfoObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.ProjectInfoObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp


        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }

    private AddPermissionsToProjectCalendarList(ListName, ProName) {

        this.ProjectCalendarObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.ProjectCalendarObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.ProjectCalendarObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.ProjectCalendarObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.ProjectCalendarObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.ProjectCalendarObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp


        }).catch(err => {
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }


    private AddPermissionsToProjectCommentHisList(ListName, ProName) {
        let reactHandler = this;

        this.ProjectCommentHisObj.breakRoleInheritance().then(res => {
            console.log("breakRoleInheritance for - ", ListName);

            this.ProjectCommentHisObj.roleAssignments.add(this.HBCAdminGrpID, 1073741829).then(res => { // 1073741829 - full controll 
                console.log(ProName, " Owners - permissions added for -", ListName);

                this.ProjectCommentHisObj.roleAssignments.add(this.ContributersGroupId, 1073741827).then(res => { // 1073741827 - Contribute
                    console.log(ProName, " Contributers - permissions added", ListName);

                    this.ProjectCommentHisObj.roleAssignments.add(this.DepartmentHeadGrpID, 1073741924).then(res => { // 1073741924 - view only
                        console.log(ProName, " View Only - permissions added", ListName);

                        this.ProjectCommentHisObj.roleAssignments.add(this.CEO_COOGrpID, 1073741924).then(res => { // 1073741924 - view only
                            console.log(ProName, " View Only - permissions added", ListName);

                            this.ProjectCommentHisObj.roleAssignments.add(this.ProjectOwnerGrpID, 1073741827).then(res => { // 1073741924 - Contribute
                                console.log(ProName, " View Only - permissions added", ListName);

                                reactHandler.cloneListItems(ProName);

                            }).catch(err => {
                                this.setState({ isLoading: false });
                                console.log("Error while adding permissions - ProjectOwnerGrp");
                            });// ProjectOwnerGrp

                        }).catch(err => {
                            this.setState({ isLoading: false });
                            console.log("Error while adding permissions - CEO_COOGrp ");
                        });// CEO_COOGrp

                    }).catch(err => {
                        this.setState({ isLoading: false });
                        console.log("Error while adding permissions - DepartmentHeadGrp");
                    });// DepartmentHeadGrp

                }).catch(err => {
                    this.setState({ isLoading: false });
                    console.log("Error while adding permissions ", ProName, " Contributers ");
                });// ContributersGroup

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while adding permissions -  HBCAdminGrp");
            });; // HBCAdminGrp

        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("Error while breakRoleInheritance for list - ", ListName)
        });
    }

    // Add permission to list end 


    // Clone list items
    private cloneListItems(ProName) {
        let flag = false;
        let oldProject = (this.state.fields['project']).split(' ').join('_');
        if (this.state.fields['clonecalender'] === true) {
            flag = true;
            this.getCalenderItems(oldProject + this.ProjectCal, ProName + this.ProjectCal);
        }
        if (this.state.fields['clonerequirements'] === true) {
            flag = true;
            this.getRequirementItems(oldProject + this.Requirements, ProName + this.Requirements);
        }
        if (this.state.fields['cloneschedule'] === true) {
            flag = true;
            this.getScheduleItems(oldProject + this.ScheduleList, ProName + this.ScheduleList);
        }
        if (this.state.fields['clonedocuments'] === true) {
            flag = true;
            this.copyDocumentListItems(oldProject + this.ProjectDocument, ProName + this.ProjectDocument);
        }
        if (!flag) {
            this.setState({ isLoading: false });
            if (this.props.id) {
                this._closePanel();
                this.props.parentMethod();
            } else {
                this._closePanel();
                this._showModal();
            }
        }
    }

    // get Calender list items
    private getCalenderItems(oldCalendarList, newCalendarList) {
        let reactHandler = this;
        let Title;
        let Description;
        let EventDate;
        let EndDate;
        let Category;
        let ParticipantsPickerId = new Array();
        let ParticipantsPicker;
        let Location;
        sp.web.lists.getByTitle(oldCalendarList).items
            .select("Title", "Description", "EventDate", "EndDate", "Category", "ParticipantsPicker/ID", "ParticipantsPicker/Title", "Location")
            .expand("ParticipantsPicker")
            .get()
            .then(res => {
                console.log("CalendarList -", res);
                for (var i = 0; i < res.length; i++) {
                    Title = res[i].Title;
                    Description = res[i].Description;
                    EventDate = res[i].EventDate;
                    EndDate = res[i].EndDate;
                    Category = res[i].Category;
                    Location = res[i].Location;
                    ParticipantsPicker = res[i].ParticipantsPicker;
                    for (var i = 0; i < ParticipantsPicker.length; i++) {
                        var ID = ParticipantsPicker[i].ID;
                        ParticipantsPickerId.push(ID);
                    }
                }
                reactHandler.addCalendarListItems(newCalendarList, Title, Description, EventDate, EndDate, Category, ParticipantsPickerId, Location);

            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("error while fetching items in ", oldCalendarList, " - ", err);
            });
    }

    // Add Calender list items
    private addCalendarListItems(newCalendarList, Title, Description, EventDate, EndDate, Category, ParticipantsPickerId, Location) {
        sp.web.lists.getByTitle(newCalendarList).items.add({
            Title: Title,
            Description: Description,
            EventDate: EventDate,
            EndDate: EndDate,
            Category: Category,
            ParticipantsPickerId: {
                results: ParticipantsPickerId  // allows multiple lookup value
            },
            Location: Location

        }).then((iar: ItemAddResult) => {
            console.log("Items added successfully in list ", newCalendarList);
            this.setState({ isLoading: false });
            if (this.props.id) {
                this._closePanel();
                this.props.parentMethod();
            } else {
                this._closePanel();
                this._showModal();
            }
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("error while adding items in ", newCalendarList, " - ", err);
        });
    }

    // get requirement list items
    private getRequirementItems(oldRequirementsList, newRequirementsList) {
        let reactHandler = this;
        let Requirement;
        let Efforts;
        let ImpactOnTimelines;
        let Resources;
        let Attachments;
        let ApproverId;
        let ApporvalStatus;
        sp.web.lists.getByTitle(oldRequirementsList).items
            .select("Requirement", "Efforts", "Impact_x0020_On_x0020_Timelines", "Resources", "Attachments",
            "Approver/ID", "Approver/Title", "Apporval_x0020_Status")
            .expand("Approver")
            .get()
            .then(res => {
                console.log("RequirementsList -", res);
                for (var i = 0; i < res.length; i++) {
                    Requirement = res[i].Requirement;
                    Efforts = res[i].Efforts;
                    ImpactOnTimelines = res[i].Impact_x0020_On_x0020_Timelines;
                    Resources = res[i].Resources;
                    Attachments = res[i].Attachments;
                    ApproverId = res[i].Approver.ID;
                    ApporvalStatus = res[i].ApporvalStatus;

                }
                this.addRequirementsListItems(newRequirementsList, Requirement, Efforts, ImpactOnTimelines, Resources, Attachments, ApproverId, ApporvalStatus);

            }).catch((err) => {
                this.setState({ isLoading: false });
                console.log("error while fetching items in ", oldRequirementsList, " - ", err);
            });
    }

    // Add requirement list items
    private addRequirementsListItems(newRequirementsList, Requirement, Efforts, ImpactOnTimelines, Resources, Attachments, ApproverId, ApporvalStatus) {
        sp.web.lists.getByTitle(newRequirementsList).items.add({
            Title: 'No Title',
            Requirement: Requirement,
            Efforts: Efforts,
            Impact_x0020_On_x0020_Timelines: ImpactOnTimelines,
            Resources: Resources,
            Attachments: Attachments,
            ApproverId: ApproverId,
            Apporval_x0020_Status: ApporvalStatus,
        }).then((iar: ItemAddResult) => {
            this.setState({ isLoading: false });
            if (this.props.id) {
                this._closePanel();
                this.props.parentMethod();
                //this.props.parentReopen();
            } else {
                this._closePanel();
                this._showModal();
            }
            console.log("Items added successfully in list ", newRequirementsList);
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("error while adding items in ", newRequirementsList, " - ", err);
        });
    }



    // get schedule list items
    private getScheduleItems(oldScheduleList, newScheduleList) {
        let reactHandler = this;
        let StartDate;
        let DueDate;
        let Duration;
        let Status0Id;
        let Priority;
        let Body;
        let Comment;
        let TaskStatus;
        let AssignedTo;
        let AssignedToId = new Array();

        sp.web.lists.getByTitle(oldScheduleList).items
            .select("StartDate", "DueDate", "Duration", "AssignedTo/Title", "AssignedTo/ID", "Status0/ID", "Status0/Status",
            "Status0/Status_x0020_Color", "Priority", "Body", "Predecessors/ID", "Predecessors/Title", "Comment", "Status")
            .expand("AssignedTo", "Status0", "Predecessors")
            .get()
            .then(res => {
                console.log("ScheduleList -", res);
                for (var i = 0; i < res.length; i++) {
                    StartDate = res[i].StartDate;
                    DueDate = res[i].DueDate;
                    Duration = res[i].Duration;
                    Status0Id = res[i].Status0.ID;
                    Priority = res[i].Priority;
                    Body = res[i].Body;
                    Comment = res[i].Comment;
                    TaskStatus = res[i].Status;

                    AssignedTo = res[i].AssignedTo;
                    for (var i = 0; i < AssignedTo.length; i++) {
                        var ID = AssignedTo[i].ID;
                        AssignedToId.push(ID);
                    }


                }

                reactHandler.addScheduleListItems(newScheduleList, StartDate, DueDate, Duration, AssignedToId, Status0Id, Priority, Body, Comment, TaskStatus);

            }).catch((err) => {
                this.setState({ isLoading: false });
                console.log("error while fetching items in ", oldScheduleList, " - ", err);
            });
    }

    // add schedule list items
    private addScheduleListItems(newScheduleList, StartDate, DueDate, Duration, AssignedToId, Status0Id, Priority, Body, Comment, TaskStatus) {
        sp.web.lists.getByTitle(newScheduleList).items.add({
            Title: 'No Title',
            StartDate: StartDate,
            DueDate: DueDate,
            Duration: Duration,
            AssignedToId: { results: AssignedToId },
            Status0Id: Status0Id,
            Priority: Priority,
            Body: Body,
            Comment: Comment,
            Status: TaskStatus

        }).then((iar: ItemAddResult) => {
            this.setState({ isLoading: false });
            console.log("Items added successfully in list ", newScheduleList);
            if (this.props.id) {
                this._closePanel();
                this.props.parentMethod();
                //this.props.parentReopen();
            } else {
                this._closePanel();
                this._showModal();
            }
        }).catch(err => {
            this.setState({ isLoading: false });
            console.log("error while adding items in ", newScheduleList, " - ", err);
        });;
    }

    // Add Document list items
    private copyDocumentListItems(oldDocumentList, newDocumentList) {



        // var siteurl = this.context.pageContext.web.absoluteUrl.split('.com')[1] + '/';
        var siteurl = '/' + window.location.pathname.split('/')[1] + '/' + window.location.pathname.split('/')[2] + '/';

        var sourceList = encodeURI(oldDocumentList) + '/';
        var destinationList = encodeURI(newDocumentList) + '/';

        sp.web.lists.getByTitle(oldDocumentList).items
            .select('Title', 'LinkFilename', '*')
            .getAll()
            .then((response) => {
                for (var i = 0; i < response.length; i++) {
                    var item = response[i];
                    sp.web.getFileByServerRelativeUrl(siteurl + sourceList + response[i].LinkFilename)
                        .copyTo(siteurl + destinationList + response[i].LinkFilename, true)
                }
                console.log(response);
                this.setState({ isLoading: false });
                if (this.props.id) {
                    this._closePanel();
                    this.props.parentMethod();
                    //this.props.parentReopen();
                } else {
                    this._closePanel();
                    this._showModal();
                }
            }).catch(err => {
                this.setState({ isLoading: false });
                console.log("Error while copying document in list ", newDocumentList, " - ", err);
            });
    }

    // private cloneListItems(ProName) {

    //     //if (this.state.isCalender === true) {
    //     //    this.getCalenderItems(this.state.CloneProject + " Project Calender", ProName + " Project Calender");
    //     //}
    //     //if (this.state.isRequirement === true) {
    //     //    this.getRequirementItems(this.state.CloneProject + " Requirements",ProName + " Requirements");
    //     //}
    //     //if (this.state.isSchedule === true) {
    //     //    this.getScheduleItems(this.state.CloneProject + " Schedule List",ProName + " Schedule List");
    //     //}
    //     //if (this.state.isDocument === true) {
    //     //    this.copyDocumentListItems(this.state.CloneProject + " Project Document",ProName + " Project Document");
    //     //}
    //     let flag = false;
    //     if (this.state.fields['clonecalender'] === true) {
    //         flag = true;
    //         this.getCalenderItems(this.state.fields['project'] + " Project Calender", ProName + " Project Calender");
    //     }
    //     if (this.state.fields['clonerequirements'] === true) {
    //         flag = true;
    //         this.getRequirementItems(this.state.fields['project'] + " Requirements", ProName + " Requirements");
    //     }
    //     if (this.state.fields['cloneschedule'] === true) {
    //         flag = true;
    //         this.getScheduleItems(this.state.fields['project'] + " Schedule List", ProName + " Schedule List");
    //     }
    //     if (this.state.fields['clonedocuments'] === true) {
    //         flag = true;
    //         this.copyDocumentListItems(this.state.fields['project'] + " Project Document", ProName + " Project Document");
    //     }
    //     if (!flag) {
    //         if (this.props.id) {
    //             this._closePanel();
    //             this.props.parentMethod();
    //             //this.props.parentReopen();
    //         } else {
    //             this._closePanel();
    //             this._showModal();
    //         }
    //     }
    // }
    // // get Calender list items
    // private getCalenderItems(oldCalendarList, newCalendarList) {
    //     let reactHandler = this;
    //     let Title;
    //     let Description;
    //     let EventDate;
    //     let EndDate;
    //     let Category;
    //     let ParticipantsPickerId = new Array();
    //     let ParticipantsPicker;
    //     let Location;
    //     sp.web.lists.getByTitle(oldCalendarList).items
    //         .select("Title", "Description", "EventDate", "EndDate", "Category", "ParticipantsPicker/ID", "ParticipantsPicker/Title", "Location")
    //         .expand("ParticipantsPicker")
    //         .get()
    //         .then(res => {
    //             console.log("CalendarList -", res);
    //             for (var i = 0; i < res.length; i++) {
    //                 Title = res[i].Title;
    //                 Description = res[i].Description;
    //                 EventDate = res[i].EventDate;
    //                 EndDate = res[i].EndDate;
    //                 Category = res[i].Category;
    //                 Location = res[i].Location;
    //                 ParticipantsPicker = res[i].ParticipantsPicker;
    //                 for (var i = 0; i < ParticipantsPicker.length; i++) {
    //                     var ID = ParticipantsPicker[i].ID;
    //                     ParticipantsPickerId.push(ID);
    //                 }
    //             }
    //             reactHandler.addCalendarListItems(newCalendarList, Title, Description, EventDate, EndDate, Category, ParticipantsPickerId, Location);

    //         }).catch(err => {
    //             console.log("error while fetching items in ", oldCalendarList, " - ", err);
    //         });
    // }

    // // Add Calender list items
    // private addCalendarListItems(newCalendarList, Title, Description, EventDate, EndDate, Category, ParticipantsPickerId, Location) {
    //     sp.web.lists.getByTitle(newCalendarList).items.add({
    //         Title: Title,
    //         Description: Description,
    //         EventDate: EventDate,
    //         EndDate: EndDate,
    //         Category: Category,
    //         ParticipantsPickerId: {
    //             results: ParticipantsPickerId  // allows multiple lookup value
    //         },
    //         Location: Location

    //     }).then((iar: ItemAddResult) => {
    //         console.log("Items added successfully in list ", newCalendarList);
    //         if (this.props.id) {
    //             this._closePanel();
    //             this.props.parentMethod();
    //             //this.props.parentReopen();
    //         } else {
    //             this._closePanel();
    //             this._showModal();
    //         }
    //     }).catch(err => {
    //         console.log("error while adding items in ", newCalendarList, " - ", err);
    //     });
    // }

    // // get requirement list items
    // private getRequirementItems(oldRequirementsList, newRequirementsList) {
    //     let reactHandler = this;
    //     let Requirement;
    //     let Efforts;
    //     let ImpactOnTimelines;
    //     let Resources;
    //     let Attachments;
    //     let ApproverId;
    //     let ApporvalStatus;
    //     sp.web.lists.getByTitle(oldRequirementsList).items
    //         .select("Requirement", "Efforts", "Impact_x0020_On_x0020_Timelines", "Resources", "Attachments",
    //         "Approver/ID", "Approver/Title", "Apporval_x0020_Status")
    //         .expand("Approver")
    //         .get()
    //         .then(res => {
    //             console.log("RequirementsList -", res);
    //             for (var i = 0; i < res.length; i++) {
    //                 Requirement = res[i].Requirement;
    //                 Efforts = res[i].Efforts;
    //                 ImpactOnTimelines = res[i].Impact_x0020_On_x0020_Timelines;
    //                 Resources = res[i].Resources;
    //                 Attachments = res[i].Attachments;
    //                 ApproverId = res[i].Approver.ID;
    //                 ApporvalStatus = res[i].ApporvalStatus;

    //             }
    //             this.addRequirementsListItems(newRequirementsList, Requirement, Efforts, ImpactOnTimelines, Resources, Attachments, ApproverId, ApporvalStatus);

    //         }).catch((err) => {
    //             console.log("error while fetching items in ", oldRequirementsList, " - ", err);
    //         });
    // }

    // // Add requirement list items
    // private addRequirementsListItems(newRequirementsList, Requirement, Efforts, ImpactOnTimelines, Resources, Attachments, ApproverId, ApporvalStatus) {
    //     sp.web.lists.getByTitle(newRequirementsList).items.add({
    //         Title: 'No Title',
    //         Requirement: Requirement,
    //         Efforts: Efforts,
    //         Impact_x0020_On_x0020_Timelines: ImpactOnTimelines,
    //         Resources: Resources,
    //         Attachments: Attachments,
    //         ApproverId: ApproverId,
    //         Apporval_x0020_Status: ApporvalStatus,
    //     }).then((iar: ItemAddResult) => {
    //         if (this.props.id) {
    //             this._closePanel();
    //             this.props.parentMethod();
    //             //this.props.parentReopen();
    //         } else {
    //             this._closePanel();
    //             this._showModal();
    //         }
    //         console.log("Items added successfully in list ", newRequirementsList);
    //     }).catch(err => {
    //         console.log("error while adding items in ", newRequirementsList, " - ", err);
    //     });
    // }



    // // get schedule list items
    // private getScheduleItems(oldScheduleList, newScheduleList) {
    //     let reactHandler = this;
    //     let StartDate;
    //     let DueDate;
    //     let Duration;
    //     let Status0Id;
    //     let Priority;
    //     let Body;
    //     let Comment;
    //     let TaskStatus;
    //     let AssignedTo;
    //     let AssignedToId = new Array();

    //     sp.web.lists.getByTitle(oldScheduleList).items
    //         .select("StartDate", "DueDate", "Duration", "AssignedTo/Title", "AssignedTo/ID", "Status0/ID", "Status0/Status",
    //         "Status0/Status_x0020_Color", "Priority", "Body", "Predecessors/ID", "Predecessors/Title", "Comment", "Status")
    //         .expand("AssignedTo", "Status0", "Predecessors")
    //         .get()
    //         .then(res => {
    //             console.log("ScheduleList -", res);
    //             for (var i = 0; i < res.length; i++) {
    //                 StartDate = res[i].StartDate;
    //                 DueDate = res[i].DueDate;
    //                 Duration = res[i].Duration;
    //                 Status0Id = res[i].Status0.ID;
    //                 Priority = res[i].Priority;
    //                 Body = res[i].Body;
    //                 Comment = res[i].Comment;
    //                 TaskStatus = res[i].Status;

    //                 AssignedTo = res[i].AssignedTo;
    //                 for (var i = 0; i < AssignedTo.length; i++) {
    //                     var ID = AssignedTo[i].ID;
    //                     AssignedToId.push(ID);
    //                 }


    //             }

    //             reactHandler.addScheduleListItems(newScheduleList, StartDate, DueDate, Duration, AssignedToId, Status0Id, Priority, Body, Comment, TaskStatus);

    //         }).catch((err) => {
    //             console.log("error while fetching items in ", oldScheduleList, " - ", err);
    //         });
    // }

    // // add schedule list items
    // private addScheduleListItems(newScheduleList, StartDate, DueDate, Duration, AssignedToId, Status0Id, Priority, Body, Comment, TaskStatus) {
    //     sp.web.lists.getByTitle(newScheduleList).items.add({
    //         Title: 'No Title',
    //         StartDate: StartDate,
    //         DueDate: DueDate,
    //         Duration: Duration,
    //         AssignedToId: { results: AssignedToId },
    //         Status0Id: Status0Id,
    //         Priority: Priority,
    //         Body: Body,
    //         Comment: Comment,
    //         Status: TaskStatus

    //     }).then((iar: ItemAddResult) => {
    //         console.log("Items added successfully in list ", newScheduleList);
    //         if (this.props.id) {
    //             this._closePanel();
    //             this.props.parentMethod();
    //             //this.props.parentReopen();
    //         } else {
    //             this._closePanel();
    //             this._showModal();
    //         }
    //     }).catch(err => {
    //         console.log("error while adding items in ", newScheduleList, " - ", err);
    //     });;
    // }

    // // Add Document list items
    // private copyDocumentListItems(oldDocumentList, newDocumentList) {

    //     var siteurl = this.context.pageContext.web.absoluteUrl.split('.com')[1] + '/';
    //     var sourceList = encodeURI(oldDocumentList) + '/';
    //     var destinationList = encodeURI(newDocumentList) + '/';

    //     sp.web.lists.getByTitle(oldDocumentList).items
    //         .select('Title', 'LinkFilename', '*')
    //         .getAll()
    //         .then((response) => {
    //             for (var i = 0; i < response.length; i++) {
    //                 var item = response[i];
    //                 sp.web.getFileByServerRelativeUrl(siteurl + sourceList + response[i].LinkFilename)
    //                     .copyTo(siteurl + destinationList + response[i].LinkFilename, true)
    //             }
    //             console.log(response);
    //             if (this.props.id) {
    //                 this._closePanel();
    //                 this.props.parentMethod();
    //                 //this.props.parentReopen();
    //             } else {
    //                 this._closePanel();
    //                 this._showModal();
    //             }
    //         }).catch(err => {
    //             console.log("Error while copying document in list ", newDocumentList, " - ", err);
    //         });
    // }






    // end
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
            }).catch(err => {
                console.log("Error in addProjectTagByTagName -", err);
            });
    }
    addProjectTag(tagName, projectID) {
        sp.web.lists.getByTitle("Project Tags").items.add({
            ProjectsId: { results: [projectID] },
            Tag: tagName
        }).then((response) => {
            // if (this.props.id) {
            //     this._closePanel();
            //     this.props.parentMethod();
            //     //this.props.parentReopen();
            // } else {
            //     this._closePanel();
            //     this._showModal();
            // }
            console.log('Project team members added -', response);
        }).catch(err => {
            console.log("Error in addProjectTag -", err);
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
                    // if (this.props.id) {
                    //     this._closePanel();
                    //     this.props.parentMethod();
                    //     //this.props.parentReopen();
                    // } else {
                    //     this._closePanel();
                    //     this._showModal();
                    // }
                    console.log(JSON.stringify(result));
                });
            }
        }).catch(err => {
            console.log("Error in updateProjectTag -", err);
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
                    <label>On Hold Status</label>
                    <select ref="status" className={formControl + " " + (this.state.errorClass["status"] ? this.state.errorClass["status"] : '')}
                        onChange={this.handleChange.bind(this, "status")} value={this.state.fields["status"]}>
                        <option>On Hold</option>
                        <option>Resume </option>
                    </select>
                </div>
            </div> : null;
        const statusDate = (this.state.showStatusDate) ?
            <div className="col-lg-6">
                <div className="form-group">
                    <span className="error">* </span><label>On Hold Date</label>
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
                    <span className="error">* </span><label>Select Project</label>
                    <select className="form-control" ref="project" onChange={this.handleChange.bind(this, "project")} value={this.state.fields["project"]}>
                        <option value="" selected disabled>Select</option>
                        {this.state.projectList.map((obj) =>
                            <option key={obj.Project} value={obj.Project}>{obj.Project}</option>
                        )}
                    </select>
                    <span className="error">{this.state.errors["project"]}</span>
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
        const attachmentDiv = (this.state.fields['projectoutline'] && this.state.fields['projectoutline'].length > 0) ?
            <div className="col-lg-6">
                {this.state.fields['projectoutline'].map((obj, i) =>
                    <div className="form-group">
                        <label style={{ float: 'left', width: '90%' }}><a href={obj.ServerRelativeUrl}><i
                            style={{ marginRight: "5px" }}
                            className='fa fa-file' ></i>{obj.FileName}</a></label>
                        <i className="far fa-times-circle" style={{ float: 'right', cursor: 'pointer' }} onClick={this.removeAttachment.bind(this, i)}></i>
                    </div>
                )}
            </div> : null;
        const emptyDiv = this.state.fields['projectoutline'] && this.state.fields['projectoutline'].length > 0 ?
            <div className="col-lg-6">
            </div> : null;
        const departmentContent = this.state.showDepartment ?
            <div className="col-sm-6 col-12">
                <div className="form-group">
                    <label>Department Name</label>
                    <select ref="departmentname" className={formControl + " " + (this.state.errorClass["departmentname"] ? this.state.errorClass["departmentname"] : '')}
                        onChange={this.handleChange.bind(this, "departmentname")} value={this.state.fields["departmentname"]}>
                        {this.state.departmentList.map((obj) =>
                            <option key={obj.Department} value={obj.ID}>{obj.Department}</option>
                        )}
                    </select>
                </div>
            </div> : null;
        return (
            // className="PanelContainer"
            !this.state.isLoading ?
                <div>

                    <Panel
                        isOpen={this.state.showPanel}
                        onDismiss={this._closePanel}
                        type={PanelType.medium}

                    >
                        <div className="PanelContainer">
                            <section className="main-content-section">
                                <div className="row">
                                    <div className="col-sm-12 col-12">
                                        <h3 className="hbc-form-header">Project Details</h3>
                                        <form name="projectform" className="hbc-form" onSubmit={this.projectSubmit.bind(this)}>
                                            <div className="row addSection">
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <span className="error">* </span><label>Project Name</label>
                                                        <input ref="projectname" type="text" className={formControl + " " + (this.state.errorClass["projectname"] ? this.state.errorClass["projectname"] : '')} placeholder="Enter project name"
                                                            onChange={this.handleChange.bind(this, "projectname")} value={this.state.fields["projectname"]} onBlur={this.handleBlurOnProjectName}>
                                                        </input>
                                                        <span className="error">{this.state.errors["projectname"]}</span>
                                                    </div>
                                                </div>
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <label>Owner</label>
                                                        <span className="calendar-style">
                                                            {this._renderControlledPicker()}
                                                        </span>
                                                        <span className="error">{this.state.errors["ownername"]}</span>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="row addSection">
                                                <div className="col-sm-12 col-12">
                                                    <div className="form-group">
                                                        {/* <label>Clone Project</label> */}
                                                        <div>
                                                            <Checkbox label="Clone Project" checked={this.state.fields["cloneproject"]} onChange={this.handleChange.bind(this, "cloneproject")} value={this.state.fields["cloneproject"]} />
                                                        </div>
                                                    </div>
                                                </div>

                                                {selectProjectContent}
                                                {chechbox1Content}
                                                {chechbox2Content}
                                                {chechbox3Content}
                                                {chechbox4Content}
                                            </div>
                                            <div className="row addSection">
                                                <div className="col-sm-12 col-12">
                                                    <div className="form-group">
                                                        <label>Project Description</label>
                                                        <textarea ref="projectdescription" style={{ height: '50px !important' }} className={formControl + " " + (this.state.errorClass["projectdescription"] ? this.state.errorClass["projectdescription"] : '')} placeholder="Brief the owner about the project"
                                                            onChange={this.handleChange.bind(this, "projectdescription")} value={this.state.fields["projectdescription"]}></textarea>
                                                        <span className="error">{this.state.errors["projectdescription"]}</span>
                                                    </div>
                                                </div>
                                                <div className="col-sm-6 col-12">
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
                                                <div className="col-sm-6 col-12">
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
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <label>Project Status</label>
                                                        <select ref="projectstatus" className={formControl + " " + (this.state.errorClass["projectstatus"] ? this.state.errorClass["projectstatus"] : '')}
                                                            onChange={this.handleChange.bind(this, "projectstatus")} value={this.state.fields["projectstatus"]}>
                                                            {this.state.statusList.map((obj) =>
                                                                <option key={obj.Status} value={obj.Id}>{obj.Status}</option>
                                                            )}
                                                        </select>
                                                    </div>
                                                </div>
                                                <div className="col-sm-6 col-12">
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
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <label>Tags</label>
                                                        <CreatableSelect
                                                            isMulti
                                                            onChange={this.handleChange2}
                                                            options={this.state.tagOptions}
                                                            value={this.state.fields["tags"]}
                                                        />
                                                        <span className="error">{this.state.errors["tags"]}</span>
                                                    </div>
                                                </div>
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <label>Risk</label>
                                                        <select className={formControl + " " + (this.state.errorClass["risk"] ? this.state.errorClass["risk"] : '')} ref="risk" onChange={this.handleChange.bind(this, "risk")} value={this.state.fields["risk"]}>
                                                            <option>Low</option>
                                                            <option>Medium</option>
                                                            <option>High</option>
                                                        </select>
                                                    </div>
                                                </div>
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <label>Project Type</label>
                                                        <div>
                                                            <Checkbox checked={this.state.fields["departmentspecific"]} label="Department Specific" onChange={this.handleChange.bind(this, "departmentspecific")} value={this.state.fields["departmentspecific"]} />
                                                        </div>
                                                    </div>
                                                </div>
                                                {/* department ID */}
                                                {departmentContent}
                                            </div>
                                            <div className="row addSection">
                                                {/* {statusContent} */}
                                                <div className="col-lg-6">
                                                    <div className="form-group">
                                                        <label>On Hold Status</label>
                                                        {/* <select ref="status" className={formControl + " " + (this.state.errorClass["status"] ? this.state.errorClass["status"] : '')}
                                                        onChange={this.handleChange.bind(this, "status")} value={this.state.fields["status"]}>
                                                        <option>On Hold</option>
                                                        <option>Resume </option>
                                                    </select> */}
                                                        <Toggle
                                                            checked={this.state.fields["status"]}
                                                            defaultChecked={false}
                                                            onText="On"
                                                            offText="Off"
                                                            onChanged={this.handleChange.bind(this, "status")}
                                                            onFocus={() => console.log('onFocus called')}
                                                            onBlur={() => console.log('onBlur called')}
                                                        />
                                                    </div>
                                                </div>
                                                {statusDate}
                                            </div>
                                            <div className="row addSection">
                                                <div className="col-sm-6 col-12">
                                                    <div className="form-group">
                                                        <label>Reccuring Project?</label>
                                                        <div>
                                                            <Checkbox label="Yes" checked={this.state.fields["requringproject"]} onChange={this.handleChange.bind(this, "requringproject")} value={this.state.fields["requringproject"]} />
                                                        </div>
                                                    </div>
                                                </div>
                                                <div className="col-sm-6 col-12">
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
                                            </div>
                                            <div className="row addSection">
                                                <div className="col-lg-6">
                                                    <div className="form-group">
                                                        <label>Project Outline</label>
                                                        <div className="fileupload" data-provides="fileupload">
                                                            <input ref="projectoutline" type="file" id="uploadFile"
                                                                onChange={this.handleChange.bind(this, "projectoutline")} >
                                                            </input>
                                                        </div>
                                                    </div>
                                                </div>
                                                {attachmentDiv}
                                            </div>
                                            <div className="row addSection">
                                                <div className="col-sm-12 col-12">
                                                    <div className="btn-sec">
                                                        <button id="submit" value="Submit" className="btn-style btn btn-success">{this.props.id ? 'Update' : 'Save'}</button>
                                                        <button type="button" className="btn-style btn btn-default" onClick={this._closePanel}>Cancel</button>
                                                    </div>
                                                </div>
                                            </div>
                                        </form>
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
                            <Link to={`/viewProjectDetails/${this.state.savedProjectID}`}>
                                <Button>Continue</Button>
                            </Link>

                        </Modal.Footer>
                    </Modal>
                </div >
                : <div style={{ textAlign: "center", fontSize: "25px" }}><i className="fas fa-spinner"></i></div>
        );
    }

    /* Api Call*/
    private getProjectByID(id): void {
        // get Project Documents list items for all projects
        let filterString = "ID eq " + id;
        sp.web.lists.getByTitle("Project").items
            .select("Project", "StartDate", "DueDate", "Risks", "Status0/ID", "Department/ID", "Department/Title", "Status0/Status", "Status0/Status_x0020_Color", "AssignedTo/Title", "AssignedTo/EMail", "AssignedTo/ID", "Priority", "Clone_x0020_Project", "Clone_x0020_Calender", "Clone_x0020_Documents", "Clone_x0020_Requirements", "Clone_x0020_Schedule", "Body", "Occurance", "Recurring_x0020_Project", "ProTypeDeptSpecific", "On_x0020_Hold_x0020_Date", "On_x0020_Hold_x0020_Status", "AttachmentFiles", "AttachmentFiles/ServerRelativeUrl", "AttachmentFiles/FileName")
            .expand("Status0", "AssignedTo", "AttachmentFiles", "Department")

            .filter(filterString)
            .getAll()
            .then((response) => {
                console.log('Project by name', response);
                this.state.fields["project"] = response ? response[0].Project : '';
                this.state.fields["projectname"] = response ? response[0].Project : '';
                this.state.fields["priority"] = response ? response[0].Priority : '';
                this.state.fields["duedate"] = response[0].DueDate ? new Date(response[0].DueDate) : null;
                this.state.fields["projectdescription"] = response ? response[0].Body : '';
                this.state.fields["startdate"] = response[0].StartDate ? new Date(response[0].StartDate) : null;
                this.state.fields["departmentspecific"] = response ? response[0].ProTypeDeptSpecific : false;
                this.state.fields["requringproject"] = response ? response[0].Recurring_x0020_Project : false;
                this.state.fields["occurance"] = response ? response[0].Occurance : '';
                this.state.fields["cloneschedule"] = response ? response[0].Clone_x0020_Schedule : false;
                this.state.fields["clonedocuments"] = response ? response[0].Clone_x0020_Documents : false;
                this.state.fields["clonerequirements"] = response ? response[0].Clone_x0020_Requirements : false;
                this.state.fields["clonecalender"] = response ? response[0].Clone_x0020_Calender : false;
                this.state.fields["cloneproject"] = response ? response[0].Clone_x0020_Project : false;
                this.state.fields["status"] = response ? response[0].On_x0020_Hold_x0020_Status : false;
                this.state.fields["projectstatus"] = response[0].Status0 ? response[0].Status0.ID : '1';
                this.state.fields["risk"] = response ? response[0].Risks : 'Low';
                this.state.fields["projectoutline"] = response ? response[0].AttachmentFiles : [];

                const selectedPeopleList: IPersonaWithMenu[] = [];
                const selectedTarget: IPersonaWithMenu = {};
                let tempSelectedPersona = {};
                if (response[0].AssignedTo && response[0].AssignedTo.length > 0) {
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

                if (response[0].ProTypeDeptSpecific && response[0].Department) {
                    this.state.fields["departmentname"] = response[0].Department.ID;
                    this.setState({ showDepartment: true });
                }

                if (response[0].Clone_x0020_Project) {
                    this.setState({ cloneProjectChecked: true });
                } else {
                    this.setState({ cloneProjectChecked: false });
                }
                if (response[0].On_x0020_Hold_x0020_Status) {
                    this.state.fields["statusdate"] = response[0].On_x0020_Hold_x0020_Date ? new Date(response[0].On_x0020_Hold_x0020_Date) : null;
                    this.setState({ showStatusDate: true });
                } else if (response[0].On_x0020_Hold_x0020_Status === null) {
                    this.state.fields["status"] = false;
                    this.setState({ showStatusDate: false });
                } else {
                    this.state.fields["statusdate"] = '';
                    this.setState({ showStatusDate: false });
                }
                console.log('State........', this.state.fields)
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
    getDepartmentList() {
        sp.web.lists.getByTitle('Departments').items.orderBy("Department", true)
            .select("ID", "Department", "Department_x0020_Owner/ID", "Department_x0020_Owner/Title").expand("Department_x0020_Owner")
            .get()
            .then(result => {
                console.log("Department - ", result);
                this.setState({ departmentList: result });
            }).catch(err => {
                console.log("Error while fetching Department - ", err);
            });
    }
    getStatusList() {
        sp.web.lists.getByTitle('Project Status Color').items
            .select('Sequence', 'Status', 'Status_x0020_Color', 'ID')
            .orderBy("Sequence")
            .get()
            .then((response: any[]) => {
                console.log("All Colors -", response);
                this.setState({ statusList: response });
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
                //fields["projectname"] = response ? response[0].Project : '';
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

                const selectedPeopleList: IPersonaWithMenu[] = [];
                const selectedTarget: IPersonaWithMenu = {};
                let tempSelectedPersona = {};
                if (response[0].AssignedTo.length > 0) {
                    response[0].AssignedTo.forEach(element => {
                        tempSelectedPersona = {
                            key: element.ID,
                            text: element.Title
                        }
                        selectedPeopleList.push(tempSelectedPersona);
                    });
                }
                this.state.fields["ownername"] = selectedPeopleList;

                this.setState({ fields, currentSelectedItems: selectedPeopleList });
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
                    itemLimit={1}
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
