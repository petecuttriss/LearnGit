import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsByActivityWebPartStrings';
import ActivityTeamsContainer from './components/TeamsByActivityContainer/TeamsByActivityContainer';
import { ITeamsByActivityContainerProps } from './components/TeamsByActivityContainer/ITeamsByActivityContainerProps';
import { setup as pnpSetup, stringIsNullOrEmpty } from '@pnp/common';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { sp, Web, Site } from '@pnp/sp';
import { IListItem } from '../../models/IListItem';
import { ListService } from '../../services/ListService';
import { IListService } from '../../models/IListService';

/* 
* 1. Use PnPJs to get the property bag items
* 2. Create a Mock to return the property bag items
* 3. Use PnPJs to get the site registry items where the Activity is equal to any of the property bag filters
* 4. Create a Mock to return the registry items
* 5. Add unit tests
* 6. Add properties
*     Title
*     Description
*     Site Registry URL
*     Central COnfiguration URL
*/

export interface ITeamsByActivityWebPartProps {
  title: string;
  description: string;
  siteUrl: string;
  listId: string;
  odataFilter: string;
  usePropertyBagFilter: boolean;
  propertyBagFilterName: string;
  hideTitle: boolean;
  hideSearchBox: boolean;
  openLinksInANewTab: boolean;
}

export default class TeamsByActivityWebPart extends BaseClientSideWebPart<ITeamsByActivityWebPartProps> {
  
  private listServiceInstance: IListService;
  
  public onInit(): Promise<void> {

    /*pnpSetup({
      spfxContext: this.context
    });*/
    this.listServiceInstance = this.context.serviceScope.consume(ListService.serviceKey);

    this._openPropertyPane = this._openPropertyPane.bind(this);

    return Promise.resolve();
  }

  public render(): void {
    let odataFilter: string = "";
    if (!this.properties.usePropertyBagFilter) {
      odataFilter = this.properties.odataFilter;
      this.renderElement(odataFilter);
    } else {
      if (this.properties.propertyBagFilterName) {
        this._getPropertyBagItem(this.properties.propertyBagFilterName).then((filter) =>{
          this.renderElement(filter);
        });
      }
      this.renderElement(odataFilter);
    }
  }

  private renderElement(odataFilter: string){

    const element: React.ReactElement<ITeamsByActivityContainerProps> = React.createElement(
      ActivityTeamsContainer,
      {
        title: this.properties.title,
        siteUrl: this.properties.siteUrl === "/" ? this.context.pageContext.site.serverRelativeUrl : this.properties.siteUrl,
        listId: this.properties.listId,
        odataFilter: odataFilter,
        hideTitle: this.properties.hideTitle,
        hideSearchBox: this.properties.hideSearchBox,
        openLinksInANewTab: this.properties.openLinksInANewTab,
        displayMode: this.displayMode,
        listService: this.listServiceInstance,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        configureStartCallback: this._openPropertyPane
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SitesRegistryUrl,
                  description: strings.SitesRegistryUrlDescription,
                  placeholder: strings.SitesRegistryUrlPlaceholder,
                  onGetErrorMessage: this._validateSiteUrl.bind(this),
                  validateOnFocusOut: true
                }),
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: this._validateListFields.bind(this),
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  webAbsoluteUrl: this.properties.siteUrl === "/" ? this.context.pageContext.site.absoluteUrl : this.properties.siteUrl
                }),
                PropertyPaneTextField('odataFilter', {
                  label: strings.ODataFilterLabel,
                  description: strings.ODataFilterDescription,
                  placeholder: strings.ODataFilterPlaceholder,
                  disabled: this.properties.usePropertyBagFilter,
                  validateOnFocusOut: true
                }),
                PropertyPaneToggle('usePropertyBagFilter', {
                  label: strings.UsePropertyBagFilterFieldLabel,
                  onText: "Yes"
                }),
                PropertyPaneTextField('propertyBagFilterName', {
                  label: strings.PropertyBagFilterLabel,
                  description: strings.PropertyBagFilterNameDescription,
                  placeholder: strings.PropertyBagFilterNamePlaceholder,
                  disabled: !this.properties.usePropertyBagFilter,
                  onGetErrorMessage: this._validatePropertyBagFilter.bind(this),
                  validateOnFocusOut: true
                }),
                PropertyPaneToggle('hideTitle', {
                  label: strings.HideTitleFieldLabel,
                  onText: "Yes"
                }),
                PropertyPaneToggle('hideSearchBox', {
                  label: strings.HideSearchBoxFieldLabel,
                  onText: "Yes"
                }),
                PropertyPaneToggle('openLinksInANewTab', {
                  label: strings.OpenLinksInANewTabFieldLabel,
                  onText: "Yes"
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _validateSiteUrl(value: string): Promise<string> {
    if (Environment.type === EnvironmentType.Local) {
      return new Promise<string>((resolve) => {
        resolve('');
        return;
      });
    }

    return new Promise<string>((resolve: (validationErrorMessage: string) => void, reject: (error: any) => void): void => {
      if (value === null ||
        value.length === 0) {
        resolve('Provide the Site Url');
        return;
      }

      let site = new Site(value);
      site.select("Title").get()
        .then((siteFound): void => {
          resolve('');
          return;
        })
        .catch((error: any): void => {
          // If it fails because previously configured web/list isn't accessible for current user
          if (error.status === 403) {
            // Still resolve with accessDenied=true
            resolve("Access Denied");
          }

          // If it fails because previously configured web/list doesn't exist anymore
          else if (error.status === 404) {
            // Still resolve with site not found
            resolve("Site Not Found");
          }

          // If it fails for any other reason, reject with the error message
          else {
            let errorMessage: string = error.statusText ? error.statusText : error;
            reject(errorMessage);
          }
        });
    });
  }

  private _validateListFields(value: string): Promise<string> {
    if (Environment.type === EnvironmentType.Local) {
      return new Promise<string>((resolve) => {
        resolve('');
        return;
      });
    }

    return new Promise<string>((resolve: (validationErrorMessage: string) => void, reject: (error: any) => void): void => {
      if (value === null ||
        value.length === 0 || value === "NO_LIST_SELECTED") {
        resolve('Select a registry list');
        return;
      }

      let web = new Web(this.properties.siteUrl);
      let requiredFields: string[] = ["Function", "Activity", "Subactivity", "SFLink", "SIAStatus", "SFDesc"];
      let missingFields: string[] = [];
      let finalArray = [];
      requiredFields.forEach((fieldName) => {
        finalArray.push(web.lists.getById(value).fields.getByInternalNameOrTitle(fieldName).get()
          .then((field): void => {
          })
          .catch((error: any): void => {
            missingFields.push(fieldName);
          }));
      });

      const result = Promise.all(finalArray).then(() => {
        if (missingFields.length > 0) {
          resolve(`The following fields are missing. Select a registry list: ${missingFields.toString()}`);
          return;
        } else {
          resolve("");
          return;
        }
      });
    });
  }

  private _validatePropertyBagFilter(value: string): Promise<string> {
    if (Environment.type === EnvironmentType.Local) {
      return new Promise<string>((resolve) => {
        resolve('');
        return;
      });
    }

    return new Promise<string>((resolve: (validationErrorMessage: string) => void, reject: (error: any) => void): void => {
      
      if(!this.properties.usePropertyBagFilter){
        resolve('');
        return; 
      }
      
      if (value === null ||
        value.length === 0) {
        if (this.properties.usePropertyBagFilter) {
          resolve('Provide the Property Bag Filter Field Name');
          return;
        }
      }

      this.listServiceInstance.getPropertyBagItem(value)
        .then((propertyFound): void => {
          if (stringIsNullOrEmpty(propertyFound)) {
            resolve('Property Value Not Found or Empty');
            return;
          } else {
            resolve('');
            return;
          }
        })
        .catch((error: any): void => {
          // If it fails because previously configured web/list isn't accessible for current user
          if (error.status === 403) {
            // Still resolve with accessDenied=true
            resolve("Access Denied");
          }

          // If it fails because previously configured web/list doesn't exist anymore
          else if (error.status === 404) {
            // Still resolve with site not found
            resolve("Site Not Found");
          }

          // If it fails for any other reason, reject with the error message
          else {
            let errorMessage: string = error.statusText ? error.statusText : error;
            reject(errorMessage);
          }
        });
    });
  }

  private _openPropertyPane(): void {
    this.context.propertyPane.open();
  }

  private _getPropertyBagItem(propertyBagIdentifier: string): Promise<string> {
    return Promise.resolve(this.listServiceInstance.getPropertyBagItem(propertyBagIdentifier).then((value) => {
      return value;
    }));
  }
}