import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IAddEventProps } from './IAddEventProps';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import "bootstrap/dist/css/bootstrap.min.css";
import { Button, Modal } from 'react-bootstrap';
import ProjectListTable from '../ProjectList/ProjectListTable';
const textcolor = {
            color: 'red' as 'red',
          }


export default class AddProject extends React.Component<IAddEventProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass:{},
    cloneProjectChecked: boolean,
    showModal: boolean;
    isDataSaved: boolean;
}> {

    constructor(props) {
        super(props);
        this.state = {
            showPanel: true,
            fields: {},
            errors: {},
            errorClass:{},
            cloneProjectChecked: false,
            showModal: false,
            isDataSaved: false
        };
        this._showModal = this._showModal.bind(this);
        this._closeModal = this._closeModal.bind(this);
    }

    handleChange(field, e) {
        if (field === 'startdate') {
            let fields = this.state.fields;
            fields[field] = e;
        }
        else if (field === 'enddate') {
            let fields = this.state.fields;
            fields[field] = e;
        }
        else {
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
        if (!fields["startdate"]) {
            formIsValid = false;
            errors["startdate"] = "Cannot be empty";
            errorClass["startdate"] = "classError";
        }
        if (!fields["enddate"]) {
            formIsValid = false;
            errors["enddate"] = "Cannot be empty";
            errorClass["enddate"] = "classError";
        }
        if (fields["startdate"] && fields["enddate"]) {
            if (fields["enddate"] < fields["startdate"]) {
                formIsValid = false;
                errors["enddate"] = "End Date should always be greater than Start Date";
                errorClass["enddate"] = "classError";
            }
        }
        // if (!fields["description"]) {
        //     formIsValid = false;
        //     errors["description"] = "Cannot be empty";
        //     errorClass["description"] = "classError";
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
            .select("ID","Title", "EndDate", "EventDate","Description" )
           
            .filter(filterString)
            .get()
            .then((response) => {
                let fields = this.state.fields;
                console.log('Project1 by name', response);
                console.log('Project112 by name',  response[0].Requirement );
                fields["projectname"] = response ? response[0].Title : '';
                fields["startdate"] =  response ? new Date(response[0].EventDate) : '';
                //fields["description"] = response ? response[0].Description : '';
                fields["enddate"] = response ? new Date(response[0].EndDate) : '';
               console.log("hieee",this.state.fields["projectname"]);
               this.setState(fields);
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
            sp.web.lists.getByTitle(this.props.list).items.getById(this.props.id).update({
                Title: obj.projectname ? obj.projectname : '',
               // Description: obj.description ? obj.description : '',
                EventDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
                //EventDate: obj.projectdescription ? obj.projectdescription : '',
                EndDate: obj.enddate ? new Date(obj.enddate).toDateString() : '',
               // Attachments: obj.filedescription ? obj.effortdescription : '',
                //Owner: obj.ownername?obj.ownername:'',
               // Impact: obj.priority ? obj.priority : '',
               // Mitigation: obj.projectdescription ? obj.projectdescription : '',
                //Department_x0020_Specific: obj.departmentspecific ? (obj.departmentspecific === 'on' ? true : false) : null,
                //Recurring_x0020_Project: obj.requringproject ? (obj.requringproject === 'on' ? true : false) : null,
                //Occurance: obj.occurance ? obj.occurance : '',
                //DepartmentId: 2,
                //Status0Id: 2

            }).then(i => {
                this._closePanel();
                this.props.parentMethod();
                this.props.parentReopen();
            });
        } else {
            sp.web.lists.getByTitle(this.props.list).items.add({
                Title: obj.projectname ? obj.projectname : '',
                EventDate: obj.startdate ? new Date(obj.startdate).toDateString() : '',
               // Description: obj.description ? obj.description : '',
                EndDate: obj.enddate ? new Date(obj.enddate).toDateString(): '',
            }).then((response) => {
                console.log('Item adding-', response);
                this.setState({ isDataSaved: true });
                this._closePanel();
                this._showModal();
                this.props.parentMethod();
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
    };

    public render(): React.ReactElement<IAddEventProps> {
        const html = '<div>Example HTML string</div>';
        let formControl = 'form-control';
        let paddingInputStyle = 'padding-input-style';
        // const selectProjectContent = this.state.cloneProjectChecked ?
        //     <div className="col-lg-12">
        //         <div className="form-group">
        //             <label>Select Project</label>
        //             <select className="form-control" ref="project" onChange={this.handleChange.bind(this, "project")} value={this.state.fields["project"]}>
        //                 <option>Project 1</option>
        //                 <option>Project 2</option>
        //                 <option>Project 3</option>
        //             </select>
        //         </div>
        //     </div> : null;

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
                                                    <h3 className="hbc-form-header">Events</h3>
                                                    {/* <div > */}
                                                        <form name="projectform" className="hbc-form" onSubmit={this.projectSubmit.bind(this)}>
                                                            <div className="row addSection">
                                                                

                                                                

                                                                <div className="col-sm-6 col-12">
                                                                    <div className="form-group">
                                                                        <label>Title <span style={textcolor}>*</span></label>
                                                                        <input ref="projectname" type="text" className={formControl + " " + (this.state.errorClass["projectname"] ? this.state.errorClass["projectname"] : '')} 
                                                                            onChange={this.handleChange.bind(this, "projectname")} value={this.state.fields["projectname"]}>
                                                                        </input>
                                                                        <span className="error">{this.state.errors["projectname"]}</span>
                                                                    </div>
                                                                </div>
                                                                {/* <div className="col-lg-6">
                                                                    <div className="form-group">
                                                                        <label>Description</label>
                                                                        <span className="calendar-style"><i className="fas fa-user icon-style"></i>
                                                                            <input ref="description" type="text" className={paddingInputStyle + " " + formControl + " " + (this.state.errorClass["description"] ? this.state.errorClass["description"] : '')}  placeholder="Brief the owner about the project"
                                                                                onChange={this.handleChange.bind(this, "description")} value={this.state.fields["description"]} dangerouslySetInnerHTML={{__html: html}}>
                                                                            </input>
                                                                        </span>
                                                                        <span className="error">{this.state.errors["ownername"]}</span>
                                                                    </div>
                                                                </div> */}
                                                                <div className="col-sm-6 col-12">
                                                                    <div className="form-group">
                                                                        <label>Impact<span style={textcolor}>*</span></label>
                                                                        <select className={formControl + " " + (this.state.errorClass["priority"] ? this.state.errorClass["priority"] : '')} ref="priority" onChange={this.handleChange.bind(this, "priority")} value={this.state.fields["priority"]}>
                                                                            <option>Meeting</option>
                                                                            <option>Business</option>
                                                                            <option>Holiday</option>
                                                                        </select>
                                                                        <span className="error">{this.state.errors["priority"]}</span>
                                                                    </div>
                                                                </div>
                                                                <div className="col-sm-6 col-12">
                                                                    <div className="form-group">
                                                                        <label>Start Time<span style={textcolor}>*</span> </label>
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
                                                                        <label>End Time<span style={textcolor}>*</span></label>
                                                                        <DatePicker
                                                                            placeholder="Select end date"
                                                                            onSelectDate={this.handleChange.bind(this, "enddate")}
                                                                            value={this.state.fields["enddate"]}
                                                                        />
                                                                        <span className="error">{this.state.errors["enddate"]}</span>
                                                                    </div>
                                                                </div>
                                                               
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

                {/* <Modal
                    show={this.state.showModal}
                    onHide={this._closeModal}
                    container={this}
                    aria-labelledby="contained-modal-title"
                    animation={false}
                > */}
                    {/* <Modal.Header>
                        <Modal.Title id="contained-modal-title">
                            Risk Created
                        </Modal.Title>
                    </Modal.Header> */}
                    {/* <Modal.Body>
                        Project Created Successfully! Do you want to configure Project
Schedule and Project Team now?
                    </Modal.Body>
                    <Modal.Footer>
                        <Button onClick={this._closeModal}>I'll Do it Later</Button>
                        <Button onClick={this._closeModal}>Continue</Button>
                    </Modal.Footer> */}
                {/* </Modal> */}
            </div>

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
