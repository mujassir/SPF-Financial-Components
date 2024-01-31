import * as React from 'react';
import * as ReactDom from 'react-dom';

import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';

import * as strings from 'ExpenseSummaryWebPartStrings';
import ExpenseSummary from './components/ExpenseSummary';
import { IExpenseSummaryProps } from './components/IExpenseSummaryProps';

export interface IExpenseSummaryWebPartProps { }

export default class ExpenseSummaryWebPart extends BaseClientSideWebPart<IExpenseSummaryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IExpenseSummaryProps> = React.createElement(
      ExpenseSummary,
      {}
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

              ]
            }
          ]
        }
      ]
    };
  }
}

