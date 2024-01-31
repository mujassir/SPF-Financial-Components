import * as React from 'react';
import * as ReactDom from 'react-dom';

import * as strings from 'MonthWiseExpensesWebPartStrings';
import MonthWiseExpenses from './components/MonthWiseExpenses';
import { IMonthWiseExpensesProps } from './components/IMonthWiseExpensesProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';
import { PropertyPaneTextField, IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IMonthWiseExpensesWebPartProps {
  listUrl: string;
  lists: any[];
}

export default class MonthWiseExpensesWebPart extends BaseClientSideWebPart<IMonthWiseExpensesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMonthWiseExpensesProps> = React.createElement(
      MonthWiseExpenses,
      {
        context: this.context,
        listUrl: this.properties.listUrl,
        lists: this.properties.lists,
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
                PropertyPaneTextField('listUrl', {
                  label: 'Site URL',
                  placeholder: 'Enter the site URL',
                  value: this.properties.listUrl,
                  deferredValidationTime: 10
                }),
                PropertyFieldListPicker('lists', {
                  label: 'Select a list',
                  selectedList: this.properties.lists,
                  disabled: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  includeListTitleAndUrl: true,
                  key: 'listPickerFieldId',
                  webAbsoluteUrl: this.properties.listUrl,
                  includeHidden: false, // Optionally include hidden lists
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
