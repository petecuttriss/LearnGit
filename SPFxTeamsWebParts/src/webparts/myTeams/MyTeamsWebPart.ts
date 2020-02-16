import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MyTeamsWebPartStrings';
import MyTeams from './components/MyTeams';
import { IMyTeamsProps } from './components/IMyTeamsProps';
import { ITeamsService } from '../../models/ITeamsService';
import { TeamsService } from '../../services/TeamsService';

export interface IMyTeamsWebPartProps {
  title: string;
  description: string;
}

export default class MyTeamsWebPart extends BaseClientSideWebPart<IMyTeamsWebPartProps> {
  private teamsServiceInstance: ITeamsService;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      this.teamsServiceInstance = this.context.serviceScope.consume(TeamsService.serviceKey);
    });
  }
  
  public render(): void {
    const element: React.ReactElement<IMyTeamsProps > = React.createElement(
      MyTeams,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        teamsService: this.teamsServiceInstance,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
