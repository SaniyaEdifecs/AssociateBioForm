import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'AssociateBioFormWebPartStrings';
import AssociateBioForm from './components/AssociateBioForm';
import { IAssociateBioFormProps } from './components/IAssociateBioFormProps';
import { sp } from "@pnp/sp"; 


export interface IAssociateBioFormWebPartProps {
  description: string;
}

export default class AssociateBioFormWebPart extends BaseClientSideWebPart<IAssociateBioFormWebPartProps> {


  public onInit(): Promise<void> {
    
    console.log("onInIt called");
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
    
    });
  
    });
    
  }
  public render(): void {
    console.log("render called");
    const element: React.ReactElement<IAssociateBioFormProps > = React.createElement(
      AssociateBioForm,
      {
        description: this.properties.description,
        context: this.context,
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
