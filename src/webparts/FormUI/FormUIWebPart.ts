import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import * as strings from 'FormUIWebPartStrings';
import FormUI from './components/FormUI';
import { IFormUIProps } from './components/IFormUIProps';

export interface IFormUIWebPartProps {
  listPageName: string;
}

export default class FormUIWebPart extends BaseClientSideWebPart<IFormUIWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFormUIProps > = React.createElement(
      FormUI,
      {
        currentUserLoginName: this.context.pageContext.user.loginName,
        context: this.context,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        listPageName: this.properties.listPageName,
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
                PropertyPaneTextField('listPageName', {
                  label: strings.ListPageNameFieldLabel
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
