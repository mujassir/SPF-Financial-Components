import * as React from 'react';
import * as ReactDom from 'react-dom';

import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, WebPartContext } from '@microsoft/sp-webpart-base';

import * as strings from 'ExpenseSummaryWebPartStrings';
import ExpenseSummary from './components/ExpenseSummary';
import { IExpenseSummaryProps } from './components/IExpenseSummaryProps';
import { IGroupByField } from './models/IGroupByField';
import { IColumnReturnProperty, IPropertyFieldRenderOption, PropertyFieldColumnPicker, PropertyFieldColumnPickerOrderBy, PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls';
import { CustomCollectionFieldType, PropertyFieldCollectionData } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldOrder } from '@pnp/spfx-property-controls/lib/PropertyFieldOrder';

export interface IExpenseSummaryWebPartProps { 
  context: WebPartContext;
  listUrl: string;
  lists: any;
  listColumns: any[];
  orderedListColumns: any[];
  groupByFields: IGroupByField[];
}

export default class ExpenseSummaryWebPart extends BaseClientSideWebPart<IExpenseSummaryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExpenseSummaryProps> = React.createElement(
      ExpenseSummary,
      {
        context: this.context,
        lists: this.properties.lists,
        listUrl: this.properties.listUrl,
        listColumns: this.properties.listColumns,
        orderedListColumns: this.properties.orderedListColumns,
        groupByFields: this.properties.groupByFields,
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected async onInit(): Promise<void> {
    this.properties.orderedListColumns = this.properties.listColumns;
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
                PropertyFieldColumnPicker('listColumns', {
                  label: 'Select columns',
                  context: this.context,
                  selectedColumn: this.properties.listColumns,
                  listId: this.properties.lists ? this.properties.lists.id : null,
                  disabled: false,
                  orderBy: PropertyFieldColumnPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  deferredValidationTime: 0,
                  key: 'multiColumnPickerFieldId',
                  displayHiddenColumns: false,
                  columnReturnProperty: IColumnReturnProperty.Title,
                  multiSelect: true,
                  renderFieldAs: IPropertyFieldRenderOption["Multiselect Dropdown"]
                }),
                PropertyFieldCollectionData("groupByFields", {
                  key: "groupByFields",
                  label: "Group By Fields",
                  panelHeader: "Group By Field Collection",
                  manageBtnLabel: "Manage Group By Fields",
                  value: this.properties.groupByFields,
                  fields: [
                    {
                      id: "column",
                      title: "Column",
                      type: CustomCollectionFieldType.dropdown,
                      options: this.properties.listColumns ? this.properties.listColumns.map(p => { return { key: p, text: p } }) : [],
                      required: true
                    },
                    {
                      id: "sortOrder",
                      title: "Sort Order",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "ascending",
                          text: "Ascending"
                        },
                        {
                          key: "descending",
                          text: "Descending"
                        }
                      ],
                      required: true
                    },

                  ],
                  disabled: false
                }),
                PropertyFieldOrder("orderedListColumns", {
                  key: "orderedListColumns",
                  label: "Column Display Order",
                  items: this.properties.orderedListColumns,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}

