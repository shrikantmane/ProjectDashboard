import * as React from "react";
import { sp, ItemAddResult, Web } from "@pnp/sp";
import { IAddTasksProps } from './IAddTaskProps';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
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

export default class AddTask extends React.Component<IAddTasksProps, {
    showPanel: boolean;
    fields: {},
    errors: {},
    errorClass: {},
    peopleList: any[],
    currentSelectedItems?: IPersonaProps[],
    mostRecentlyUsed: IPersonaProps[];
    delayResults: boolean,
    currentPicker?: number | string,
    statusList: any,
}>{
    private _picker: IBasePicker<IPersonaProps>;

    constructor(props) {
        super(props);
        const peopleList: IPersonaWithMenu[] = [];

        this.state = {
            showPanel: true,
            fields: {},
            errors: {},
            errorClass: {},
            peopleList: peopleList,
            currentSelectedItems: [],
            mostRecentlyUsed: [],
            delayResults: false,
            currentPicker: 1,
            statusList: [],
        };

        this.handleBlurOntitle = this.handleBlurOntitle.bind(this);
    }
    componentDidMount() {
        this._getAllSiteUsers();
        this.getStatusList();
    }
    getStatusList() {
        sp.web.lists.getByTitle('Task Status Color').items
            .select('Sequence', 'Status', 'Status_x0020_Color', 'ID')
            .orderBy("Sequence")
            .get()
            .then((response: any[]) => {
                console.log("All Colors -", response);
                this.setState({ statusList: response });
            });
    }
    handleChange(field, e, isChecked: boolean) {
        if (field === 'startdate') {
            let fields = this.state.fields;
            fields[field] = e;
            this.validateDateDiff(field);
        }
        else if (field === 'duedate') {
            let fields = this.state.fields;
            fields[field] = e;
            this.validateDateDiff(field);
        }
        else {
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
        if (!fields["title"]) {
            formIsValid = false;
            errors["title"] = "Cannot be empty";
            errorClass["title"] = "classError";
        }
        if (fields["startdate"] && fields["duedate"]) {
            if (fields["duedate"] < fields["startdate"]) {
                formIsValid = false;
                errors["duedate"] = "Due Date should always be greater than Start Date";
                errorClass["duedate"] = "classError";
            }
        }

        this.setState({ errors: errors, errorClass: errorClass });
        return formIsValid;
    }
    taskSubmit(e) {
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
            this.setState({ fields });;
            sp.web.lists.getByTitle("All Tasks").items.add({
                Title: obj.title ? obj.title : '',
                StartDate: obj.startdate ? new Date(obj.startdate).toDateString() : null, //"2018-08-03T07:00:00Z"
                DueDate: obj.duedate ? new Date(obj.duedate).toDateString() : null,   //"2018-08-03T07:00:00Z" 
                Status0Id: obj.status ? obj.status : 1,
                Body: obj.description ? obj.description : '',
                AssignedToId: { results: obj.ownername },
                //ProjectId: null,
                DepartmentId: 3
            }).then((iar: ItemAddResult) => {
                console.log("Task Added !");
            }).catch(err => {
                console.log("Error while adding task", err);
            });

        } else {
            console.log('form has error');
        }
    }
    validateDateDiff(field) {
        let fields = this.state.fields;
        let errors = {};
        let errorClass = {};
        let formIsValid = true;
        if (fields["startdate"] && fields["duedate"]) {
            if (fields["duedate"] < fields["startdate"]) {
                formIsValid = false;
                if (field === 'startdate') {
                    errors["startdate"] = "Start Date should always be less than Due Date";
                    errorClass["startdate"] = "classError";
                } else {
                    errors["duedate"] = "Due Date should always be greater than Start Date";
                    errorClass["duedate"] = "classError";
                }
            } else {
                formIsValid = true;
                errors["duedate"] = "";
                errorClass["duedate"] = "";
                errors["startdate"] = "";
                errorClass["startdate"] = "";
            }
        }
        this.setState({ errors: errors, errorClass: errorClass });
        return formIsValid;
    }
    private _closePanel = (): void => {
        this.setState({ showPanel: false });
        this.props.parentReopen();
    };
    handleBlurOntitle() {
        console.log(this.state.fields['title']);
        let errors = this.state.errors;
        let errorClass = this.state.errorClass;
        if (this.state.fields['title'] && this.state.fields['title'].trim()) {
            errors["title"] = "";
            errorClass["title"] = "";
            this.setState({ errors: errors, errorClass: errorClass });
        } else {
            errors["title"] = "Cannot be empty";
            errorClass["title"] = "classError";
            this.setState({ errors: errors, errorClass: errorClass });
        }
    }
    public render(): React.ReactElement<IAddTasksProps> {
        let formControl = 'form-control';
        return (
            <div >
                <Panel
                    isOpen={this.state.showPanel}
                    onDismiss={this._closePanel}
                    type={PanelType.medium}

                >
                    <div className="PanelContainer">
                        <section className="main-content-section">
                            <div className="row">
                                <div className="col-sm-12 col-12">
                                    <h3 className="hbc-form-header">Add Task</h3>
                                    <form name="projectform" className="hbc-form" onSubmit={this.taskSubmit.bind(this)}>
                                        <div className="row addSection">
                                            <div className="col-sm-6 col-12">
                                                <div className="form-group">
                                                    <span className="error">* </span><label>Title <i className="fa fa-info-circle" aria-hidden="true" data-toggle="tooltip" title="Maximum 50 characters are allowed"></i></label>
                                                    <input ref="title" type="text" maxLength={50} className={formControl + " " + (this.state.errorClass["title"] ? this.state.errorClass["title"] : '')} placeholder="Enter task name"
                                                        onChange={this.handleChange.bind(this, "title")} value={this.state.fields["title"]} onBlur={this.handleBlurOntitle}>
                                                    </input>
                                                    <span className="error">{this.state.errors["title"]}</span>
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
                                            <div className="col-sm-12 col-12">
                                                <div className="form-group">
                                                    <label> Description</label>
                                                    <textarea ref="description" style={{ height: '50px !important' }} className={formControl + " " + (this.state.errorClass["description"] ? this.state.errorClass["description"] : '')} placeholder="Brief the owner about the task"
                                                        onChange={this.handleChange.bind(this, "description")} value={this.state.fields["description"]}></textarea>
                                                    <span className="error">{this.state.errors["description"]}</span>
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
                                                    <label>Status</label>
                                                    <select ref="status" className={formControl + " " + (this.state.errorClass["status"] ? this.state.errorClass["status"] : '')}
                                                        onChange={this.handleChange.bind(this, "status")} value={this.state.fields["status"]}>
                                                        {this.state.statusList.map((obj) =>
                                                            <option key={obj.Status} value={obj.Id}>{obj.Status}</option>
                                                        )}
                                                    </select>
                                                </div>
                                            </div>
                                        </div>
                                        <div className="row addSection">
                                            <div className="col-sm-12 col-12">
                                                <div className="btn-sec">
                                                    <button id="submit" value="Submit" className="btn-style btn btn-success">Save</button>
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
            </div >
        )
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