import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'TaskActivityWebPartStrings';
import TaskActivity from './components/TaskActivity';
import { ITaskActivityProps } from './components/ITaskActivityProps';
import { sp } from "@pnp/sp";
import { WebPartContext } from '@microsoft/sp-webpart-base';
//import "@pnp/sp/webs";
export interface ITaskActivityWebPartProps {
  description: string;
  context: WebPartContext;
}

export default class TaskActivityWebPart extends BaseClientSideWebPart<ITaskActivityWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {   
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  public render(): void {
    const element: React.ReactElement<ITaskActivityProps > = React.createElement(
      TaskActivity,
      {
        description: this.properties.description,
        context: this.context,
      }
    );
    
    ReactDom.render(element, this.domElement);
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
