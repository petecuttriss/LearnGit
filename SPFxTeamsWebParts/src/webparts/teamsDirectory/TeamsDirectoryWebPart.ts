import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsDirectoryWebPartStrings';
import TeamsDirectory from './components/TeamsDirectory';
import { ITeamsDirectoryProps } from './components/ITeamsDirectoryProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { sp, Web, Site } from '@pnp/sp';
import { IListItem } from '../../models/IListItem';
import { setup as pnpSetup, stringIsNullOrEmpty } from '@pnp/common';
import { ListService } from '../../services/ListService';
import { IListService } from '../../models/IListService';
import { MockListService } from '../../services/MockListService';

export interface ITeamsDirectoryWebPartProps {
  title: string;
  description: string;
  siteUrl: string;
  listId: string;
  hideTitle: boolean;
  hideSearchBox: boolean;
  openLinksInANewTab: boolean;
}

export default class TeamsDirectoryWebPart extends BaseClientSideWebPart<ITeamsDirectoryWebPartProps> {
  private listServiceInstance: IListService;

  public onInit(): Promise<void> {

    /*pnpSetup({
      spfxContext: this.context
    });*/
    this.listServiceInstance = this.context.serviceScope.consume(ListService.serviceKey);
    //this.listServiceInstance = new MockListService();

    this._openPropertyPane = this._openPropertyPane.bind(this);

    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<ITeamsDirectoryProps> = React.createElement(
      TeamsDirectory,
      {
        title: this.properties.title,
        siteUrl: this.properties.siteUrl === "/" ? this.context.pageContext.site.serverRelativeUrl : this.properties.siteUrl,
        listId: this.properties.listId,
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

  private _openPropertyPane(): void {
    this.context.propertyPane.open();
  }
}
