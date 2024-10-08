import * as React from 'react';
import { IInputs } from "./generated/ManifestTypes";
import "./customStyles.css";
import { IBasePicker, IBasePickerSuggestionsProps, initializeIcons, IPersonaProps, NormalPeoplePicker, ValidationState } from '@fluentui/react';


export interface IPeoplePersona {
  fullName?: string;
  email?: string;
}

export interface ITeamMemberPickerProps {
  people?: any;
  preselectedpeople?: any;
  context?: ComponentFramework.Context<IInputs>;
  peopleList?: (newValue: any) => void;
  isPickerDisabled?: boolean;
}

export interface ITeamMemberPickerState {
  currentPicker?: number | string;
  delayResults?: boolean;
  peopleList: IPersonaProps[];
  mostRecentlyUsed: IPersonaProps[];
  currentSelectedItems?: IPersonaProps[];
}

const suggestionProps: IBasePickerSuggestionsProps = {
  // suggestionsHeaderText: 'Suggested People',
  // mostRecentlyUsedHeaderText: 'Suggested Contacts',
  // noResultsFoundText: 'No results found',
  loadingText: 'Loading',
  showRemoveButtons: false,
  // suggestionsAvailableAlertText: 'People Picker Suggestions available',
  // suggestionsContainerAriaLabel: 'Suggested contacts'
};


export class TeamMemberPickerTypes extends React.Component<any, ITeamMemberPickerState> {
  // All pickers extend from BasePicker specifying the item type.
  private _picker = React.createRef<IBasePicker<IPersonaProps>>();

  constructor(props: ITeamMemberPickerProps) {
    super(props);
    const personas: IPersonaProps[] = this.props.preselectedpeople !== undefined || null ? this.props.preselectedpeople.map((user: { fullName: any; email: any; text: any; secondaryText: any; }) => ({
      text: user.fullName !== undefined || null ? user.fullName : user.text !== undefined || null ? user.text : "",
      secondaryText: user.email !== undefined || null ? user.email : user.secondaryText !== undefined || null ? user.secondaryText : "",
    })) : [];

    this.state = {
      currentPicker: 1,
      delayResults: false,
      peopleList: this.props.people,
      mostRecentlyUsed: [],
      currentSelectedItems: personas,

    };
    initializeIcons();
    this.handleChange = this.handleChange.bind(this);
  }

  public render() {

    return (
        this.props.context.parameters.selectionType.raw! === "Single" ?
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
            defaultSelectedItems={this.state.currentSelectedItems}
            onItemSelected={this._onItemSelected}
            itemLimit={1}
            inputProps={{
              onBlur: (ev: React.FocusEvent<HTMLInputElement>) => {
              },
              onFocus: (ev: React.FocusEvent<HTMLInputElement>) => {
              },
              'aria-label': 'People Picker'
            }}
            componentRef={this._picker}
            onInputChange={this._onInputChange}
            resolveDelay={300}
            disabled={this.props.isPickerDisabled}
            onChange={() => this.handleChange()}
          />
          :
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
            defaultSelectedItems={this.state.currentSelectedItems}
            onItemSelected={this._onItemSelected}
            inputProps={{
              onBlur: (ev: React.FocusEvent<HTMLInputElement>) => {
              },
              onFocus: (ev: React.FocusEvent<HTMLInputElement>) => {
              },
              'aria-label': 'People Picker'
            }}
            componentRef={this._picker}
            onInputChange={this._onInputChange}
            resolveDelay={300}
            disabled={this.props.isPickerDisabled}
            onChange={() => this.handleChange()}
          />

    );

  }

  private handleChange() {
    if (this._picker !== undefined) {
      this.getPeopleRelatedVal();
    }
  }

  private _onItemSelected = (item: any): Promise<IPersonaProps> => {
    const processedItem = { ...item };
    processedItem.text = `${item.text}`;
    return new Promise<IPersonaProps>((resolve, reject) =>
      resolve(processedItem));
  };



  private async getPeopleRelatedVal() {
    try {
      const tempPeople: IPeoplePersona[] = this._picker.current!.items!.map(item => ({
        fullName: item.text,
        email: item.secondaryText
      }));
      if (this.props.peopleList) {
        this.props.peopleList(tempPeople);
      }
    }
    catch (err) {
      console.log(err);
    }
  }

  private _getTextFromItem(persona: IPersonaProps): string {
    return persona.text as string;
  }

  private _onRemoveSuggestion = (item: IPersonaProps): void => {
    const { peopleList, mostRecentlyUsed: mruState } = this.state;
    const indexPeopleList: number = peopleList.indexOf(item);
    const indexMostRecentlyUsed: number = mruState.indexOf(item);

    if (indexPeopleList >= 0) {
      const newPeople: IPersonaProps[] = peopleList.slice(0, indexPeopleList).concat(peopleList.slice(indexPeopleList + 1));
      this.setState({ peopleList: newPeople });
    }

    if (indexMostRecentlyUsed >= 0) {
      const newSuggestedPeople: IPersonaProps[] = mruState
        .slice(0, indexMostRecentlyUsed)
        .concat(mruState.slice(indexMostRecentlyUsed + 1));
      this.setState({ mostRecentlyUsed: newSuggestedPeople });
    }
  }

  private _onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[] | undefined,
    limitResults?: number
  ): any => {
    if (filterText) {
      if (filterText.length > 2) {
        const currentPeople = currentPersonas !== undefined ? currentPersonas : [];
        return this._searchUsers(filterText, currentPeople);
      }
    } else {
      return [];
    }
  };

  private _searchUsers(filterText: string, currentPersonas: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
    return new Promise(async (resolve: any, reject: any) => {
      let People: IPersonaProps[] = [];

      try {
        const entityName = this.props.context.parameters.entityName.raw!;
        let tempPeople: any;
        if (entityName === "systemuser") {

          tempPeople = await this.props.context.webAPI.retrieveMultipleRecords(entityName, `?$select=fullname,internalemailaddress&$filter=startswith(fullname,'${filterText}')`);
          People = tempPeople.entities.map((entity: any) => ({
            text: entity.fullname,
            secondaryText: entity.internalemailaddress
          }));
        } else if (entityName === "aaduser") {
          tempPeople = await this.props.context.webAPI.retrieveMultipleRecords(entityName, `?$select=displayname,mail,accountenabled&$filter=startswith(displayname,'${filterText}') and usertype eq 'member' and mail ne null`);
          People = tempPeople.entities
            .filter((entity: any) =>
              this.props.context.parameters.loadEnabledAccounts.raw! === 'true'
                ? entity.accountenabled
                : true
            )
            .map((entity: any) => ({
              text: entity.displayname,
              secondaryText: entity.mail
            }));
        } else if (entityName === "team" && this.props.context.parameters.teamId.raw! !== "12345678-1234-1234-1234-123456789012") {
          tempPeople = await this.props.context.webAPI.retrieveMultipleRecords("teammemberships", `?$select=systemuserid&$filter=teamid eq ${this.props.context.parameters.teamId}`);
          const usersMatched = await this.props.context.webAPI.retrieveMultipleRecords(entityName, `?$select=fullname,internalemailaddress,systemuserid&$filter=startswith(fullname,'${filterText}')`);
          const matchedTeamUsers = usersMatched.entities.filter((user: any) => tempPeople.entities.some((people: any) => people.systemuserid === user.systemuserid));
          People = matchedTeamUsers.map((entity: any) => ({
            text: entity.fullname,
            secondaryText: entity.internalemailaddress
          }));
        }
        People = this._removeDuplicates(People, currentPersonas);

        resolve(People);
      } catch (err) {
        console.log(err);
        reject(People);
      }
    });
  }

  private _returnMostRecentlyUsed = (currentPersonas: any): IPersonaProps[] | Promise<IPersonaProps[]> => {
    let { mostRecentlyUsed } = this.state;
    mostRecentlyUsed = this._removeDuplicates(mostRecentlyUsed, currentPersonas);
    return this._filterPromise(mostRecentlyUsed);
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

  // private _filterPersonasByText(filterText: string): IPersonaProps[] {
  //     return this.state.peopleList.filter(item => this._doesTextStartWith(item.text as string, filterText));
  // }

  // private _doesTextStartWith(text: string, filterText: string): boolean {
  //     return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  // }

  private _convertResultsToPromise(results: IPersonaProps[]): Promise<IPersonaProps[]> {
    return new Promise<IPersonaProps[]>((resolve, reject) => setTimeout(() => resolve(results), 2000));
  }

  private _removeDuplicates(personas: IPersonaProps[], possibleDupes: IPersonaProps[]) {
    return personas.filter(persona => !this._listContainsPersona(persona, possibleDupes));
  }

  private _validateInput = (input: string): ValidationState => {
    if (input.indexOf('@') !== -1) {
      return ValidationState.valid;
    } else if (input.length > 1) {
      return ValidationState.warning;
    } else {
      return ValidationState.invalid;
    }
  }

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

}