import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TeamsTabsWebPartStrings';
import TeamsTabs from './components/TeamsTabs';
import { ITeamsTabsProps } from './components/ITeamsTabsProps';
import { ITeamsService } from '../../models/ITeamsService';
import { TeamsService } from '../../services/TeamsService';

export interface ITeamsTabsWebPartProps {
  title: string;
  description: string;
}

export default class TeamsTabsWebPart extends BaseClientSideWebPart<ITeamsTabsWebPartProps> {
  private teamsServiceInstance: ITeamsService;

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      /*sp.setup({
        spfxContext: this.context
      });

      graph.setup({
        spfxContext: this.context
      });*/
      this.teamsServiceInstance = this.context.serviceScope.consume(TeamsService.serviceKey);
    });
  }

  public render(): void {
    const element: React.ReactElement<ITeamsTabsProps> = React.createElement(
      TeamsTabs,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        teamsService: this.teamsServiceInstance,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
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
