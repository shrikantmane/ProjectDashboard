import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IAddDocumentProps } from './IAddDocumentProps';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import "bootstrap/dist/css/bootstrap.min.css";
import { Button, Modal } from 'react-bootstrap';
import ProjectListTable from '../ProjectList/ProjectListTable';
const textcolor = {
            color: 'red' as 'red',
          }


export default class AddProject extends React.Component<IAddDocumentProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass:{},
    cloneProjectChecked: boolean,
    showModal: boolean;
    isDataSaved: boolean;
    selectedFile:any,
    showDiv: boolean,
    isLoading: boolean,
    showComponent:boolean
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
            isDataSaved: false,
            selectedFile: "",
            showDiv:false,
            isLoading: false,
            showComponent:false
        };
        this._showModal = this._showModal.bind(this);
        this._closeModal = this._closeModal.bind(this);
        this.UploadFiles=this.UploadFiles.bind(this);
        this.handleBlurOnFile=this.handleBlurOnFile.bind(this);
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
    
    UploadFiles() {

        //in case of multiple files,iterate or else upload the first file.
    if(this.state.selectedFile!=""){
    this.setState({
        showComponent:false
    })

    

        var file = this.state.selectedFile;
        if (file != undefined || file != null) {
           
            
                this.setState({
                    isLoading:true ,
                   // showComponent:false
            
                })
            
            
            //assuming that the name of document library is Documents, change as per your requirement, 
            //this will add the file in root folder of the document library, if you have a folder named test, replace it as "/Documents/test"
            sp.web.getFolderByServerRelativeUrl(this.props.list).files.add(file.name, file, true).then((result) => {
                console.log(file.name + " upload successfully!");
                this.setState({
                    selectedFile :"",
                    isLoading:false
                  });
                  this._closePanel();
                  this.props.parentMethod();
                  this.props.parentReopen();
                // this.getProjectDocuments(this.props.list);
                //this.setState({ isLoading: false });
    

            });

        }
    }
      else{
        this.setState({
            showComponent:true
        })
    }
    }
    fileChangedHandler = (event) => {
        const file = event.target.files[0]
        this.setState({ selectedFile: event.target.files[0],
         })
    }
    handleBlurOnFile(){
        console.log(this.state.fields['projectname']);
            let errors = this.state.errors;
            let errorClass = this.state.errorClass;
            if (!this.state.selectedFile)
            {
                this.setState({ 
                    showComponent: true
                   
                });
            } 
            else{
                this.setState({ 
                    showComponent: false
                     });
            }
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

    public render(): React.ReactElement<IAddDocumentProps> {
        const html = '<div>Example HTML string</div>';
        let formControl = 'form-control';
        let paddingInputStyle = 'padding-input-style';
        
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
              {/* <div className="wrapper"> */}

              <div className="row">
                <div className="col-sm-12 col-12">
                  {/* <section id="step1">
                                            <div className="well">
                                                <div className="row"> */}
                  <h3 className="hbc-form-header">Documents</h3>
                  {/* <div > */}
                  <form
                    name="projectform"
                    className="hbc-form"
                    onSubmit={this.projectSubmit.bind(this)}
                  >
                    <div className="row addSection">
                      <div className="col-sm-6 col-12">
                        <div className="form-group">
                        <span className="error">* </span> <label>
                            Documents
                          </label>
                          <input
                            type="file"
                            onChange={this.fileChangedHandler}
                            onBlur={this.handleBlurOnFile}
                          />
                          <button
                            type="button"
                            className="btn-style btn btn-success"
                            style={{ marginBottom: "10px" }}
                            onClick={this.UploadFiles}
                          >
                            Upload
                          </button>
                          {this.state.showComponent ?
                        <div style={textcolor}>Please upload file</div>:
                        null
                    }
                          <span className="error">
                            {this.state.errors["projectname"]}
                          </span>
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
 : <div style={{ textAlign: "center", fontSize: "25px" }}><i className="fa fa-spinner fa-spin"></i></div>
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
