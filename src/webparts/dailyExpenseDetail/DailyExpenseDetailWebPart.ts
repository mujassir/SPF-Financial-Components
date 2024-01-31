import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DailyExpenseDetailWebPartStrings';
import DailyExpenseDetail from './components/DailyExpenseDetail';
import { IDailyExpenseDetailProps } from './components/IDailyExpenseDetailProps';

export interface IDailyExpenseDetailWebPartProps {
  description: string;
}

export default class DailyExpenseDetailWebPart extends BaseClientSideWebPart<IDailyExpenseDetailWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDailyExpenseDetailProps> = React.createElement(
      DailyExpenseDetail,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }



  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
