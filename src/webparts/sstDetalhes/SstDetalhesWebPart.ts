import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SstDetalhesWebPartStrings';
import SstDetalhes from './components/SstDetalhes';
import { ISstDetalhesProps } from './components/ISstDetalhesProps';

export interface ISstDetalhesWebPartProps {
  description: string;
  idListaProject: string;
  idListaIssues: string;
}

export default class SstDetalhesWebPart extends BaseClientSideWebPart<ISstDetalhesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISstDetalhesProps> = React.createElement(
      SstDetalhes,
      {
        description: this.properties.description,
        context: this.context,
        siteurl: this.context.pageContext.web.absoluteUrl,
        idListaProject: this.properties.idListaProject,
        idListaIssues: this.properties.idListaIssues
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
                }),
                PropertyPaneTextField('idListaProject', {
                  label: "GUID da lista Projetos"
                }),
                PropertyPaneTextField('idListaIssues', {
                  label: "GUID da lista Issues"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
