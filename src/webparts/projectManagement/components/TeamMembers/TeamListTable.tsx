import * as React from "react";
import { sp, ItemAddResult } from "@pnp/sp";
import { DataTable } from "primereact/components/datatable/DataTable";
import { Column } from "primereact/components/column/Column";
import styles from "../ProjectManagement.module.scss";
import  {ITeamMembersProps } from "./ITeamMembersProps";
import { ITeamState } from "./ITeamState";
import {
    TeamMembers
} from "./TeamList";
import { find, filter, sortBy } from "lodash";
import { SPComponentLoader } from "@microsoft/sp-loader";
import AddProject from '../AddProject/AddProject';


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
            currentSelectedItems: []
        };
        this.onAddProject = this.onAddProject.bind(this);
        this.refreshGrid = this.refreshGrid.bind(this);
    }
    refreshGrid (){
        this.getAllProjectMemeber()
    }
    dt: any;
    componentDidMount() {
        SPComponentLoader.loadCss(
            "https://use.fontawesome.com/releases/v5.1.0/css/all.css"
        );
        this.getAllProjectMemeber();
        this._getAllSiteUsers();
    }
    componentWillReceiveProps(nextProps) { }



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

    

    ownerTemplate(rowData: TeamMembers, column) {
        if (rowData.Team_x0020_Member)
            return (
                <div>
                    {rowData.Team_x0020_Member.Title}
                </div>
            );
    }
    actionTemplate(rowData, column) {
        return <a href="#"> Remove</a>;
    }
    editTemplate(rowData, column) {
        return <a href="#"> Edit </a>;
    }
    onAddProject() {
        console.log('button clicked');
        this.setState({
            showComponent: true,
        });
    }
    public render(): React.ReactElement<ITeamState> {
        return (
            <div>
                {/* <DataTableSubmenu /> */}
                {this._renderNormalPicker()}
                <div className="content-section implementation">
                    <button type="button" className="btn btn-outline btn-sm" style={{ marginBottom: "10px" }} onClick={this.onAddProject}>
                        Add Team Members 1
                    </button>
                    {/* {this.state.showComponent ?
                            <AddProject parentMethod={this.refreshGrid}/>  :
                        null
                    } */}

                    <DataTable value={this.state.projectList} paginator={true} rows={10} rowsPerPageOptions={[5, 10, 20]}>
                        <Column field="AssignedTo" header="Owner" body={this.ownerTemplate} />
                       
                        <Column field="Start_x0020_Date" header="Start Date" body={this.duedateTemplate}  />
                        <Column field="End_x0020_Date" header="End Date" body={this.enddateTemplate} />
                        <Column field="Status" header="Status" />
                       
                        <Column header="Remove" body={this.actionTemplate} />
                    </DataTable>
                </div>

                {/* <DataTableDoc></DataTableDoc> */}
            </div>
        );
    }

    /* Api Call*/

    

    getAllProjectMemeber(){
        sp.web.lists.getByTitle("Project Team Members").items.select("Team_x0020_Member/ID", "Team_x0020_Member/Title","Start_x0020_Date", "End_x0020_Date","Status").expand("Team_x0020_Member").getAll().then((response) => {
            console.log('member by name', response);
            this.setState({ projectList: response });
            
        });
      }

          /* Private Methods */

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

              imageUrl: persona.Email === undefined || persona.Email === '' ?  null : profileUrl
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
      
      private _renderNormalPicker() {
        return (
          <NormalPeoplePicker
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
            resolveDelay={300}
          />
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