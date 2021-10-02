import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FabricComboWpWebPartStrings';
import FabricComboWp from './components/FabricComboWp';
import { IFabricComboWpProps } from './components/IFabricComboWpProps';
import FabricUiComboBox from './components/FabricComboWp';

export interface IFabricComboWpWebPartProps {
  description: string;
}

export const getChoiceFields = async (webURL,field) => {
  let resultarr = [];
  alert(webURL+" "+field);
  await fetch(`${webURL}/_api/web/lists/GetByTitle('ComboBoxExample')/fields?$filter=EntityPropertyName eq '${field}'`, {
      method: 'GET',
      mode: 'cors',
      credentials: 'same-origin',
      headers: new Headers({
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'Access-Control-Allow-Origin': '*',
          'Cache-Control': 'no-cache',
          'pragma': 'no-cache',
      }),
  }).then(async (response) => await response.json())
      .then(async (data) => {
          for (var i = 0; i < data.value[0].Choices.length; i++) {
              
              await resultarr.push({
                  key:data.value[0].Choices[i],
                  text:data.value[0].Choices[i]
            });
      }
      });
  return await resultarr;
};

export default class FabricComboWpWebPart extends BaseClientSideWebPart<IFabricComboWpWebPartProps> {

  public async render(): Promise<void> {
    const element: React.ReactElement<IFabricComboWpProps> = React.createElement(
      FabricUiComboBox,
      {
        description: this.properties.description,
        webURL:this.context.pageContext.web.absoluteUrl,
        singleValueChoices: await getChoiceFields(this.context.pageContext.web.absoluteUrl,'SingleValueComboBox'),
        multiValueChoices: await getChoiceFields(this.context.pageContext.web.absoluteUrl,'MultiValueComboBox')       
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
