import * as React from 'react';
import styles from './DailyExpenseDetail.module.scss';
import type { IDailyExpenseDetailProps } from './IDailyExpenseDetailProps';
import { repeat } from 'lodash';
import { SPHttpClient } from '@microsoft/sp-http';
import { Alert, Box } from '@mui/material';
import Constants from '../common/constants';

export default class DailyExpenseDetail extends React.Component<IDailyExpenseDetailProps, {
  errorMessage: string;
  getLoader: boolean;
  listItems: any[];
  viewFields: any[];
  groupByFields: any[];
}> {
  public state;
  constructor(props: IDailyExpenseDetailProps) {
    super(props);

    this.state = {
      startDate: '',
      endDate: '',
      errorMessage: '',
      getLoader: false,
      listItems: [],
      viewFields: [],
      groupByFields: [],
    };
  }

  private async getListData(): Promise<void> {
    let listUrl = '';
    if (this.props.listUrl) listUrl = this.props.listUrl;
    if (!listUrl || listUrl.length === 0) {
      this.setState({ errorMessage: 'Please add SPFX WEB URL.' });
      return;
    }

    let listTitle = '';
    if (this.props.lists) listTitle = this.props.lists.title;
    if (!listTitle || listTitle.length === 0) {
      this.setState({ errorMessage: 'Please select a list first.' });
      return;
    }

    this.setState({ errorMessage: '' });
    this.setState({ getLoader: true });

    const startDate = new Date();
    startDate.setHours(0);
    startDate.setMinutes(0);
    startDate.setSeconds(0);
    startDate.setMilliseconds(0);

    const endDate = new Date();
    endDate.setHours(23);
    endDate.setMinutes(59);
    endDate.setSeconds(59);

    let filter = `Date ge '${startDate.toISOString()}' and Date le '${endDate.toISOString()}'`;

    try {
      // Get All Fields of the list
      const endpoint = `${listUrl}/_api/web/lists/getbytitle('${listTitle}')/fields`;
      const response = await this.props.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      const getFields = await response.json();
      const allFields = getFields.value || [];

      const titleToInternalNameMap = new Map();
      allFields.forEach((field: { Title: any; InternalName: any }) => {
        titleToInternalNameMap.set(field.Title, field.InternalName);
      });

      let complexFieldNamesArray: string[] = [];
      let internalNamesArray = this.props.listColumns.map(title => {
        return titleToInternalNameMap.get(title) || title; // Fallback to title if mapping not found
      });
      let viewFields: { name: any; displayName: any; isResizable: boolean; sorting: boolean }[] = [];

      let groupByFields: { name: string }[] = [];
      if (this.props.groupByFields?.length > 0) {
        groupByFields = this.props.groupByFields.map(d => {
          return {
            name: titleToInternalNameMap.get(d.column) || d.column,
          };
        });
      }
      if (this.props.orderedListColumns) {

        const complexFields = allFields.filter((p: any) => this.props.orderedListColumns.indexOf(p.Title) > -1 && p.FieldTypeKind === 20);
        complexFieldNamesArray = complexFields.map((p: any) => p.InternalName);
        for (let index = 0; index < internalNamesArray.length; index++) {
          if (complexFieldNamesArray.indexOf(internalNamesArray[index]) > -1)
            internalNamesArray[index] = internalNamesArray[index] + "/Title";
        }

        const groupNames: any = this.props.groupByFields?.map(e => e.column) || [];
        viewFields = this.props.orderedListColumns.filter(e => !groupNames.includes(e)).map(title => {
          return {
            name: titleToInternalNameMap.get(title) || title,
            displayName: title,
            isResizable: true,
            sorting: true,
            minWidth: 100,
            maxWidth: 100
          };
        });
      }

      // Get list items
      const listItemEndpoint = `${listUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=${internalNamesArray.join(',')}&$expand=${complexFieldNamesArray.join(',')}&$filter=${filter}&$top=${Constants.Defaults.MaxPageSize}`;
      const listItemResponse = await this.props.context.spHttpClient.get(listItemEndpoint, SPHttpClient.configurations.v1);
      const listItems = await listItemResponse.json();

      this.setState({
        viewFields: viewFields,
        groupByFields: groupByFields,
        listItems: listItems.value.map((p: any) => {
          if (complexFieldNamesArray.length > 0) {
            complexFieldNamesArray.forEach(element => {
              p[element] = p[element].Title;
            });
          }
          return p;
        }),
        errorMessage: '',
      });
    } catch (error) {
      this.setState({ errorMessage: error.message });
    } finally {
      this.setState({ getLoader: false });
    }
  }

  async componentDidMount() {
    this.getListData()
  }

  public render(): React.ReactElement<IDailyExpenseDetailProps> {
    return (
      <Box>
        {this.state.errorMessage && <Alert severity="error"> {this.state.errorMessage}</Alert>}

        <div>
          <h2>Daily Expense Detail</h2>
        </div>

        {this.renderListView()}
      </Box>
    );
  }
  public renderListView() {
    const items = this.state.listItems;
    const groupByFields = this.state.groupByFields.map((e: { name: string }) => e.name);

    const groupTree = this.createTreeView(items, groupByFields)
    const tableView = this.renderTable(groupTree)
    return (tableView)
  }

  private createTreeView(dataset: any[], groupByColumns: string[]) {
    const root = { name: 'Root', children: [] };

    dataset.forEach((record: any) => {
      let currentNode: any = root;

      groupByColumns.forEach(column => {
        const key = record[column];
        let childNode: any = currentNode.children.filter((child: any) => child.name === key)[0]

        if (!childNode) {
          childNode = { name: key, children: [] };
          currentNode.children.push(childNode);
        }
        currentNode = childNode;
      });
      currentNode.children.push(record);
    });

    return root.children;
  }

  private renderTable(data: any[]): React.ReactNode {
    const viewFields = this.state.viewFields;
    if (data.length === 0) {
      return ('Data not Exist')
    }

    return (
      <div>
        <table className={styles.strippedTable}>
          <thead>
            <tr>
              {
                viewFields.map((field: any) => (
                  <th key={field.name}>{field.displayName}</th>
                ))
              }
            </tr>
          </thead>
          <tbody>
            {this.renderTableRows(data)}
          </tbody>
        </table>
      </div>
    );
  }

  private renderTableRows(data: any[], level: number = 0): React.ReactNode {
    return data.map((item: any, index: number) => {
      let levelClass
      switch (level) {
        case 0:
          levelClass = styles.listLevel0
          break;
        default:
          levelClass = styles.listLevel1
          break;

      }
      // Calculate the sum of debit and credit for each parent
      let totalAmount = 0;
      if (item?.children) {
        item.children.forEach((child: any) => {
          totalAmount += (child.Amount || 0) + (child.Credit || 0);
        });
        return (
          <>
            <tr>
              <td colSpan={100} className={`${levelClass} ${styles.parentRow}`}>
                {repeat('-- ', level)}
                {`${item.name || ''}`}
                {totalAmount > 0 ? ` (${totalAmount.toFixed(2)})` : ''}
              </td>
            </tr>
            {this.renderTableRows(item.children, level + 1)}
          </>
        );
      } else {
        const viewFields = this.state.viewFields;
        return (
          <tr key={index}>
            {
              viewFields.map((field: any) => (
                <td key={field.name}>
                  {(() => {
                    switch (field.name) {
                      case 'Date':
                        return new Intl.DateTimeFormat('en-US').format(new Date(item[field.name]));
                      default:
                        return item[field.name];
                    }
                  })()}
                </td>
              ))
            }
          </tr>
        );
      }
    });
  }
}
