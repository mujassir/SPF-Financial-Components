import * as React from 'react';
import type { IExpenseSummaryProps } from './IExpenseSummaryProps';
import { repeat } from 'lodash';
import { SPHttpClient } from '@microsoft/sp-http';
import Constants from '../common/constants';
import { Alert, Box, Grid, TextField } from '@mui/material';
import { LoadingButton } from '@mui/lab';
import styles from './ExpenseSummary.module.scss';
import { entries } from 'lodash';

export default class ExpenseSummary extends React.Component<IExpenseSummaryProps, {
  startDate: string;
  endDate: string;
  errorMessage: string;
  getLoader: boolean;
  listItems: any[];
  viewFields: any[];
  titleToInternalNameMap: Map<string, string>;
  groupByFields: any[];
}> {
  public state;
  constructor(props: IExpenseSummaryProps) {
    super(props);

    this.state = {
      startDate: '',
      endDate: '',
      errorMessage: '',
      getLoader: false,
      listItems: [],
      viewFields: [],
      titleToInternalNameMap: new Map(),
      groupByFields: [],
    };
  }

  async componentDidMount() {
    const today = new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth(), 1); // First day of this month
    const endDate = today.toISOString(); // Today's date
    this.setState({ startDate: startDate.toISOString(), endDate: endDate });

    setTimeout(() => {
      this.getListData()
    }, 200);
  }

  private handleDateChange = (date: Date | string, field: 'startDate' | 'endDate'): void => {
    if (typeof date === 'string') date = new Date(date);
    switch (field) {
      case 'startDate':
        this.setState({ startDate: date ? date.toISOString() : '' });
        break;
      case 'endDate':
        this.setState({ endDate: date ? date.toISOString() : '' });
        break;
    }
  };

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

    const { startDate, endDate } = this.state;

    let filter = ``;
    if (startDate) {
      if (filter) filter += ' and ';
      filter += `Date ge '${startDate}'`;
    }
    if (endDate) {
      if (filter) filter += ' and ';
      const nextDay = new Date(endDate);
      nextDay.setDate(nextDay.getDate() + 1);
      filter += `Date le '${nextDay.toISOString()}'`;
    }

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
        titleToInternalNameMap,
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
  public render(): React.ReactElement<IExpenseSummaryProps> {
    const { startDate, endDate } = this.state
    return (
      <Box>
        {this.state.errorMessage && <Alert severity="error"> {this.state.errorMessage}</Alert>}

        <div>
          <h2>Expense Summary
            &nbsp; <small>
              ({startDate && (new Date(startDate)).toDateString()}
              &nbsp; to &nbsp;
              {endDate && (new Date(endDate)).toDateString()})
            </small>
          </h2>
        </div>

        {this.renderFilter()}
        {this.renderListView()}
      </Box>
    );
  }

  public renderFilter(): React.ReactElement<IExpenseSummaryProps> {
    return (
      <Grid container spacing={1}>
        <Grid item>
          {this.renderDateTextField('Start Date', 'startDate')}
        </Grid>
        <Grid item>
          {this.renderDateTextField('End Date', 'endDate')}
        </Grid>
        <Grid item>
          <LoadingButton onClick={() => this.getListData()} loading={this.state.getLoader} loadingPosition="end" variant='contained'>
            <span>Get Data</span>
          </LoadingButton>
        </Grid>
      </Grid>
    );
  }

  private renderDateTextField(label: string, field: 'startDate' | 'endDate'): React.ReactNode {
    return (
      <TextField
        type='date'
        label={label}
        variant="outlined"
        size="small"
        focused
        onChange={(event: React.ChangeEvent<HTMLInputElement>) => this.handleDateChange(event.target.value, field)}
      />
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
        const amount = dataset.filter(e => e[column] === key).map(e => e.Debit).reduce((partialSum, a) => partialSum + a, 0)?.toFixed(2)

        if (!childNode) {
          childNode = { name: key, parent: column, amount, children: [] };
          currentNode.children.push(childNode);
        }
        currentNode = childNode;
      });

      currentNode.children.push(record);
    });

    return root.children;
  }

  private renderTable(data: any[]): React.ReactNode {
    if (data.length === 0) {
      return ('Data not Exist')
    }

    return (
      <div>
        <table className={styles.strippedTable}>
          <thead>
            <tr>
              <th style={{ width: '100%', textAlign: 'left' }}>Description</th>
              <th>Amount</th>
            </tr>
          </thead>
          <tbody>
            {this.renderTableRows(data)}
          </tbody>
        </table>
      </div>
    );
  }

  private getGroupTitle(map: any, value: string): string | undefined {
    for (const [key, val] of entries(map)) {
      if (val === value) {
        return key;
      }
    }
    return undefined; // If value is not found
  }


  private renderTableRows(data: any[], level: number = 0): React.ReactNode {
    const titleToInternalNameMap = this.state.titleToInternalNameMap;

    return data.map((item: any,) => {
      if (item?.children) {
        const parent = this.getGroupTitle(titleToInternalNameMap, item.parent)
        return (
          <>
            <tr>
              <td>
                {repeat('-- ', level)}
                {`${parent || ''}: `}
                <strong>{item.name || ''}</strong>
              </td>
              <td style={{ textAlign: 'right' }}>
                {`${item.amount}`}
              </td>
            </tr>
            {this.renderTableRows(item.children, level + 1)}
          </>
        );
      }
      return null
    });
  }

}
