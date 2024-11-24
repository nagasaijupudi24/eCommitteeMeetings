/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'XenWpCommitteeMeetingsFormsWebPartStrings';
import XenWpCommitteeMeetingsCreateForm from './components/XenWpCommitteeMeetingsCreateForm';
// import { IXenWpCommitteeMeetingsFormsProps } from './components/IXenWpCommitteeMeetingsFormsProps';
import { spfi, SPFx } from '@pnp/sp';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';
import XenWpCommitteeMeetingsViewForm from './components/XenWpCommitteeMeetingsViewForm';

export interface IXenWpCommitteeMeetingsFormsWebPartProps {
  description: string;
  sp:any;
  listName:any;
  committeeMeetingNameList:any;
  formType:string;
  libraryId:any;
  homePageUrl:string;
  passCodeUrl:any;
}

export default class XenWpCommitteeMeetingsFormsWebPart extends BaseClientSideWebPart<IXenWpCommitteeMeetingsFormsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    console.log(this.properties)
    let element:any;
    if(this.properties.formType==="New"){
      element=React.createElement(
        XenWpCommitteeMeetingsCreateForm,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          sp:this.properties.sp,
          listName:this.properties.listName,       
          committeeMeetingNameList:this.properties.committeeMeetingNameList,
          formType:this.properties.formType,//formType
          libraryId:this.properties.libraryId,
          homePageUrl:this.properties.homePageUrl
          
        })
    }else{
      element=React.createElement(
        XenWpCommitteeMeetingsViewForm,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context: this.context,
          sp:this.properties.sp,
          listName:this.properties.listName,       
          committeeMeetingNameList:this.properties.committeeMeetingNameList,//formType
          formType:this.properties.formType,//formType
          libraryId:this.properties.libraryId,
          homePageUrl:this.properties.homePageUrl,
          passCodeUrl:this.properties.passCodeUrl
        })

    }
    

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.properties.sp=spfi().using(SPFx(this.context));
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyFieldListPicker('listName', {
                  label: 'Select a list',
                  selectedList: this.properties.listName,
                  includeHidden: false,
                  includeListTitleAndUrl:true,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange:this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  // onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  baseTemplate:100,
                  key: 'listPickerFieldId'
                }),
              
                PropertyFieldListPicker('committeeMeetingNameList', {
                  label: 'Select a Committee Name list',
                  selectedList: this.properties.committeeMeetingNameList,
                  includeHidden: false,
                  includeListTitleAndUrl:true,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange:this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  // onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  baseTemplate:100,
                  key: 'listPickerFieldId'
                }),

                PropertyFieldListPicker('libraryId', {
                  label: 'Select a Library',
                  selectedList: this.properties.libraryId,
                  includeHidden: false,
                  includeListTitleAndUrl: true,
                  orderBy: PropertyFieldListPickerOrderBy.Id,
                  disabled: false,
                  baseTemplate: 101,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  // onGetErrorMessage: null,/
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect: false,

                }),
                PropertyPaneDropdown('formType', {
                  label: "Form Type",
                  selectedKey: 'New',
                  options: [
                    { key: 'New', text: 'New' },
                    { key: 'View', text: 'View' },
                    // { key: 'Edit', text: 'Edit' },
                    // { key: 'allRequest', text: 'All Request' },

                  ]
                }),

                PropertyPaneTextField('homePageUrl', {
                  label:"ReDirect Url",
                  value: this.properties.homePageUrl
                }),

                PropertyPaneTextField('passCodeCreateUrl', {
                  label: "Create Passcode URL",
                  // Use a default value for the home URL if the description is not provided.
                  value: this.properties.passCodeUrl,
                  resizable: true,
                  // placeholder: "Enter home URL or description here"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
