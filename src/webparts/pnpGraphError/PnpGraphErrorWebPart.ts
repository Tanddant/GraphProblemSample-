import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PnpGraphErrorWebPartStrings';
import PnpGraphError from './components/PnpGraphError';
import { IPnpGraphErrorProps } from './components/IPnpGraphErrorProps';
import { graph } from "@pnp/graph/presets/all";

export interface IPnpGraphErrorWebPartProps {
  description: string;
}

export default class PnpGraphErrorWebPart extends BaseClientSideWebPart<IPnpGraphErrorWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      graph.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    /*     const element: React.ReactElement<IPnpGraphErrorProps> = React.createElement(
          PnpGraphError,
          {
            description: this.properties.description
          }
        );
    
        ReactDom.render(element, this.domElement); */


    // A simple loading message
    this.domElement.innerHTML = `Loading...`;

    // here we will load the current web's properties
    graph.groups.get().then(groups => {

      this.domElement.innerHTML = `Groups: <ul>${groups.map(g => `<li>${g.displayName}</li>`).join("")}</ul>`;
    });
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
