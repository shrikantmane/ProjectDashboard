import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import { ITeamMembersProps } from "./ITeamMembersProps";
import { ITeamState } from "./ITeamState";
import {
    TeamMembers
} from "./TeamList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';
import AddTeam from '../AddteamMembers/AddTeam';

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



export default class ProjectListTable extends React.Component<
    ITeamMembersProps,
    ITeamState
    > {
    private _picker: IBasePicker<IPersonaProps>;
    constructor(props) {
        super(props);

        const peopleList: IPersonaWithMenu[] = [];

        // this.state = {
        //    projectList: new Array<Project>(),
        //     //   projectTimeLine: new Array<ProjectTimeLine>(),
        //     projectName: null,
        //     ownerName: null,
        //     status: null,
        //     priority: null,
        //     isLoading: true,
        //     isTeamMemberLoaded: false,
        //     isKeyDocumentLoaded: false,
        //     isTagLoaded: false,
        //     expandedRowID: -1,
        //     expandedRows: []
        // };
        this.state = {
            projectList: new Array<TeamMembers>(),
            showComponent: false,
            currentPicker: 1,
            delayResults: false,
            peopleList: peopleList,
            mostRecentlyUsed: [],
            currentSelectedItems: [],
            fields: {},
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
        this.actionTemplate=this.actionTemplate.bind(this);
        this.editTemplate = this.editTemplate.bind(this);
        this.reopenPanel = this.reopenPanel.bind(this);
    }
    refreshGrid() {
        this.setState({
            showComponent: false,
            projectID: null
        })
        this.getAllProjectMemeber(this.props.list);
    }
    componentWillReceiveProps(nextProps) {
        if (nextProps.list != "" || nextProps.list != null) {
            this.getAllProjectMemeber(nextProps.list);
        }
    }

    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        if ((this.props.list) != "" || (this.props.list) != null) {
            this.getAllProjectMemeber(this.props.list);
        }
        this._getAllSiteUsers();
    }



    /* Html UI */
    duedateTemplate(rowData: TeamMembers, column) {
        if (rowData.Start_x0020_Date)
            return (
                <div>
                    {(new Date(rowData.Start_x0020_Date)).toLocaleDateString()}
                </div>
            );
    }
    enddateTemplate(rowData: TeamMembers, column) {
        if (rowData.End_x0020_Date)
            return (
                <div>
                    {(new Date(rowData.End_x0020_Date)).toLocaleDateString()}
                </div>
            );
    }
    reopenPanel() {
        this.setState({
            showComponent: false,
            projectID: null
        });
    }
    private onEditProject(rowData, e): any {
        e.preventDefault();
        console.log('Edit :' + rowData);
        this.setState({
            showComponent: true,
            projectID: rowData.ID
        });
    }


    ownerTemplate(rowData: TeamMembers, column) {
        if (rowData.Team_x0020_Member)
            return (
                <div>
                    {rowData.Team_x0020_Member.Title}
                </div>
            );
    }
    actionTemplate(rowData, column) {
        if(rowData.Status==="Active"){
        return <a href="#" onClick={this.deleteListItem.bind(this, rowData)}><i className="fas fa-user-times"></i></a>;
             }
     else{
            return<div style={{display : "none"}}></div>
        }
    }
    editTemplate(rowData, column) {
        return <a href="#" onClick={this.onEditProject.bind(this, rowData)}><i className="far fa-edit"></i></a>;
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    private deleteListItem(rowData,e):any {
        e.preventDefault();
           console.log('Edit :' + rowData);
           
           
        
           sp.web.lists.getByTitle(this.props.list).
           items.getById(rowData.ID).update({
            Status: "Inactive"
          }).then((response) => {
            console.log("Project Team Member item updated");
            this.getAllProjectMemeber(this.props.list);
          });
       
      
       }
    public render(): React.ReactElement<ITeamState> {
        return (
            <div className="">
                {/* <DataTableSubmenu /> */}
                <div className="content-section implementation">
                    <h5>Team Members</h5>
                    <div className="display-line">
                        {/* {this._renderControlledPicker()} */}
                        {/* {this.state.showComponent ?
                            <AddProject parentMethod={this.refreshGrid}/>  :
                        null
                    } */}
                        <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                            Add Members
                        </button>
                        {this.state.showComponent ?
                        <AddTeam id={this.state.projectID} parentReopen={this.reopenPanel} parentMethod={this.refreshGrid} list={this.props.list} projectId={this.props.projectId} memberlist={this.state.projectList} /> :
                        null
                    }
                    </div>
                    
                    <div className="member-list">
                        <DataTable value={this.state.projectList} responsive={true} paginator={true} rows={5} rowsPerPageOptions={[5, 10, 20]}>
                        <Column body={this.editTemplate}style={{ width: "8%", textAlign:"center" }} />
                            <Column field="AssignedTo" header="Name" sortable={true} body={this.ownerTemplate}style={{ width: "35%" }} />

                            <Column field="Start_x0020_Date" sortable={true} header="Assigned Date" body={this.duedateTemplate}style={{ width: "31%" }} />
                            {/* <Column field="End_x0020_Date" sortable={true} header="End Date" body={this.enddateTemplate}style={{ width: "21%" }} /> */}
                            <Column field="Status" sortable={true} header="Status"style={{ width: "26%" }} />

                            <Column body={this.actionTemplate}style={{ width: "7%" }} />
                        </DataTable>
                    </div>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/



    getAllProjectMemeber(list) {
        if ((list) != "") {
            sp.web.lists.getByTitle(list).items.select("ID", "Team_x0020_Member/ID", "Team_x0020_Member/Title", "Start_x0020_Date", "End_x0020_Date", "Status")
                .expand("Team_x0020_Member").get().then((response) => {
                    console.log('members by name', response);
                    this.setState({ projectList: response });

                });
        }
    }

    /* Private Methods */

    /*Start: People Picker Methods */
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
                    itemLimit={1}
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