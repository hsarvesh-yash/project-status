import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ProjectStatusWebPartStrings';
import ProjectStatus from './components/ProjectStatus';
import { IProjectStatusProps } from './components/IProjectStatusProps';

export interface IProjectStatusWebPartProps {
  backgroundImageUrl: string;
}

export default class ProjectStatusWebPart extends BaseClientSideWebPart<IProjectStatusWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IProjectStatusProps> = React.createElement(
      ProjectStatus,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        context: this.context,
        backgroundImageUrl: this.properties.backgroundImageUrl || '',
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('backgroundImageUrl', {
                  label: strings.BackgroundImageUrlFieldLabel,
                  description: 'Absolute URL to the PowerPoint slide background image.',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
