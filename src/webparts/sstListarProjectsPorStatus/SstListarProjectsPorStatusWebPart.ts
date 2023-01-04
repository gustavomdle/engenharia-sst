import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SstListarProjectsPorStatusWebPartStrings';
import SstListarProjectsPorStatus from './components/SstListarProjectsPorStatus';
import { ISstListarProjectsPorStatusProps } from './components/ISstListarProjectsPorStatusProps';

export interface ISstListarProjectsPorStatusWebPartProps {
  description: string;
}

export default class SstListarProjectsPorStatusWebPart extends BaseClientSideWebPart<ISstListarProjectsPorStatusWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISstListarProjectsPorStatusProps> = React.createElement(
      SstListarProjectsPorStatus,
      {
        description: this.properties.description
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
