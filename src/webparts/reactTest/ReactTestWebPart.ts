import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import * as strings from 'ReactTestWebPartStrings';

import ReactTest from './components/ReactTest';
import { IReactTestProps } from './components/IReactTestProps';
import { AadTokenProvider } from '@microsoft/sp-http';
import { Providers, SharePointProvider, MgtPerson} from '@microsoft/mgt';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IReactTestWebPartProps {
  description: string;
}

export default class ReactTestWebPart extends BaseClientSideWebPart<IReactTestWebPartProps> {

  public render(): void { 
    this.context.msGraphClientFactory.getClient()
    .then((client: MSGraphClient): void => {   
      const element: React.ReactElement<IReactTestProps> = React.createElement(
        ReactTest,
        {
          description: this.properties.description,
          graphClient: client
        }
      );
      ReactDom.render(element, this.domElement);
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
