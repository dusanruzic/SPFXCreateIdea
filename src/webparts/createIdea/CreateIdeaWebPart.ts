import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { sp } from "@pnp/sp";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { Environment } from "@microsoft/sp-core-library";


import * as strings from 'CreateIdeaWebPartStrings';
import CreateIdea from './components/CreateIdea';
import { ICreateIdeaProps } from './components/ICreateIdeaProps';
import SharePointService from '../../services/SharePoint/SharePointService';


export interface ICreateIdeaWebPartProps {
  description: string;
}

export default class CreateIdeaWebPart extends BaseClientSideWebPart<ICreateIdeaWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICreateIdeaProps > = React.createElement(
      CreateIdea,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(() =>{
      //let ideaListID = 'CF70FB14-EE3E-4D16-921A-3449856770E7';
      let ideaListID = 'Idea';
      SharePointService.setup(this.context, Environment.type, ideaListID);
      sp.setup({
        spfxContext: this.context
      });

  });}

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
