import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HelloWorldWebPartStrings';
import HelloWorld from './components/HelloWorld';
import { IHelloWorldProps } from './components/IHelloWorldProps';
import * as microsoftTeams from '@microsoft/teams-js';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart <IHelloWorldWebPartProps> {
  
  private _teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<void>{
    return new Promise<void>((resolve, reject)=>{
        if (this.context.microsoftTeams) {
          this.context.microsoftTeams.getContext(context =>{
            this._teamsContext = context;
            resolve();
          });
        }
        else{
          resolve();
        }
    });
  }

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        description: this.properties.description,
        teamsContext: this._teamsContext,
        webpartContext: this.context
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
