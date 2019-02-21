import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { sp } from "@pnp/sp";
import * as strings from 'InternalProjectsWebPartStrings';
import InternalProjects from './components/InternalProjects';
import { IInternalProjectsProps } from './components/IInternalProjectsProps';

export interface IInternalProjectsWebPartProps {
  description: string;
}

export default class InternalProjectsWebPart extends BaseClientSideWebPart<IInternalProjectsWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IInternalProjectsProps > = React.createElement(
      InternalProjects,
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
