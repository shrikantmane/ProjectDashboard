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


export default class AddProject extends React.Component<IAddProjectProps, {
    showPanel: boolean;
}> {

    constructor(props) {
        super(props);
        this.state = {
            showPanel: true
        };
    }

    public render(): React.ReactElement<IAddProjectProps> {
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

                                    <div className="col-md-8">
                                        <section id="step1">
                                            <div className="well">
                                                <div className="row">
                                                    <h3>Project Details</h3>
                                                    <div >
                                                        <div className="row">
                                                            <div className="col-lg-12">
                                                                <label>Clone Project</label>
                                                                <div>
                                                                    <span className="col-lg-12 col-sm-12 radBtn">
                                                                        <input type="checkbox" id="all" name="selectorAssignor">
                                                                        </input>
                                                                        {/* <label for="all"></label> */}
                                                                        <div className="check"></div>
                                                                        <p className="checkbox-title"></p>

                                                                    </span>

                                                                </div>
                                                            </div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Project Name</label>
                                                                    {/* onChange={this.handleChange.bind(this, "name")} value={this.state.fields["name"]} */}
                                                                    <input ref="name" type="text" className="form-control" placeholder="Enter project name">
                                                                    </input>

                                                                </div>
                                                            </div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Owner</label>
                                                                    <span className="calendar-style"><i className="fas fa-user icon-style"></i>
                                                                        <input type="text" className="padding-input-style form-control" placeholder="Enter owners name">
                                                                        </input>
                                                                    </span>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-12">
                                                                <div className="form-group">
                                                                    <label>Project Description</label>
                                                                    <textarea className="form-control" placeholder="Brief the owner about the project"></textarea>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Start Date</label>
                                                                    <span className="calendar-style"><i className="far fa-calendar-alt icon-style"></i>
                                                                        <input type="text" className="padding-input-style form-control datepicker" placeholder="01/01/1999">
                                                                        </input>
                                                                    </span>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>End Date</label>
                                                                    <span className="calendar-style"><i className="far fa-calendar-alt icon-style"></i>
                                                                        <input type="text" className="padding-input-style form-control datepicker" placeholder="01/01/1999">
                                                                        </input>
                                                                    </span>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Priority</label>
                                                                    <select className="form-control">
                                                                        <option>Low</option>
                                                                        <option>Medium</option>
                                                                        <option>High</option>
                                                                    </select>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Project Type</label>
                                                                    <div className="display-line">
                                                                        <span className="col-lg-12 col-sm-12 radBtn">
                                                                            <input type="checkbox" id="1" name="selectorAssignor">
                                                                            </input>
                                                                            {/* <label for="1"></label>   */}
                                                                            <div className="check"></div>
                                                                            <p className="checkbox-title">Department Specific	</p>
                                                                        </span>
                                                                    </div>

                                                                </div>
                                                            </div>

                                                            <div className="clearfix"></div>

                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Tags</label>
                                                                    <input type="text" className="form-control" placeholder="Enter Tags">
                                                                    </input>
                                                                </div>
                                                            </div>

                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Requring Project?</label>
                                                                    <div className="display-line">
                                                                        <span className="col-lg-12 col-sm-12 radBtn">
                                                                            <input type="checkbox" id="2" name="selectorAssignor">
                                                                            </input>
                                                                            {/* <label for="2"></label>  */}
                                                                            <div className="check"></div>
                                                                            <p className="checkbox-title">Yes	</p>
                                                                        </span>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                            <div className="clearfix"></div>
                                                            <div className="col-lg-6">
                                                                <div className="form-group">
                                                                    <label>Occurance</label>
                                                                    <select className="form-control">
                                                                        <option>Daily</option>
                                                                        <option>Weekly </option>
                                                                        <option>Months</option>
                                                                    </select>
                                                                </div>
                                                            </div>
                                                            <div className="col-lg-12">
                                                                <div className="text-center btn-sec">
                                                                    {/* onClick={(e) => this.contactSubmit(e)} */}
                                                                    <button type="button" className="btn btn-primary"  >Save</button>
                                                                    <button type="button" className="btn btn-default">Cancel</button>
                                                                </div>
                                                            </div>


                                                        </div>
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
            </div>

        );
    }
    private _closePanel = (): void => {
        this.setState({ showPanel: false });
    };
    /* Api Call*/

}
