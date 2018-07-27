import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { IProjectViewProps } from "./IProjectViewProps";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "../ProjectManagement.module.scss";
export default class ProjectViewDetails extends React.Component<
    IProjectViewProps
    > {
    constructor(props) {
        super(props);
        this.state = {

        };
    }
    componentDidMount() {
        const {
      match: { params }
    } = this.props;
        console.log('params : ' + params.id)
    }

    public render(): React.ReactElement<IProjectViewProps> {
        return (
            <section className="main-content-section">
                <div className="wrapper">
                    <div className="row">
                        <div className="col-lg-12 col-md-12 col-sm-12">
                            <section id="step1">
                                <div className="well">
                                    <div className="row">
                                        <div className="clearfix"></div>
                                        <div className="row portfolioNTasks">
                                            {/* first box */}
                                            <div className="col-xs-12 col-sm-12 col-md-12 col-lg-6">
                                                <div className="well recommendedProjects">
                                                    <div className="row">
                                                        <div className="col-sm-12 col-md-12 col-lg-12 cardHeading">
                                                            <div> 
                                                                <h5 className="pull-left heading-style">Team Members</h5>
                                                            </div>
                                                        </div>
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
        );
    }
}
