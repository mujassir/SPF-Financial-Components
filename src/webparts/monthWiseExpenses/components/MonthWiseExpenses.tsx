import * as React from 'react';
import type { IMonthWiseExpensesProps } from './IMonthWiseExpensesProps';
import { Alert, Box } from '@mui/material';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './MonthWiseExpenses.module.scss';
import Constants from '../common/constants';

export default class MonthWiseExpenses extends React.Component<IMonthWiseExpensesProps, {
  errorMessage: string;
  getLoader: boolean;
  monthlyItems: any[];
  accountHeads: any[];
  monthsList: any[];
}> {
  public state;
  context: any
  constructor(props: IMonthWiseExpensesProps) {
    super(props);

    this.state = {
      errorMessage: '',
      getLoader: false,
      monthlyItems: [],
      accountHeads: [],
      monthsList: [],
    };
  }

  private async getLast12MonthListData(): Promise<void> {
    let listUrl = '';
    let listTitle = '';

    if (this.props.listUrl) listUrl = this.props.listUrl;
    if (!listUrl || listUrl.length == 0) {
      this.setState({ errorMessage: 'Please add SPFX WEB URL.' });
      return;
    }

    if (this.props.lists) listTitle = this.props.lists.title;
    if (!listTitle || listTitle.length == 0) {
      this.setState({ errorMessage: 'Please select a list first.' });
      return;
    }

    try {
      const currentDate = new Date();
      const startDate = new Date(currentDate.toISOString());
      startDate.setMonth(currentDate.getMonth() - 12);
      startDate.setDate(1);
      const endDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + 2, 0);

      const monthlyItemsEndpoint = `${listUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$filter=Date ge '${startDate.toISOString()}' and Date le '${endDate.toISOString()}'&$top=${Constants.Defaults.MaxPageSize}`;
      const monthlyItemsResponse = await this.props.context.spHttpClient.get(monthlyItemsEndpoint, SPHttpClient.configurations.v1);
      const monthlyItems = await monthlyItemsResponse.json();
      if (monthlyItems?.error) throw new Error(monthlyItems.error.message);

      const monthsList: any = this.getMonthsWithStartEndDates(startDate, endDate)
      const accountHeads: any = monthlyItems.value.map((e: any) => e.AccountHead);
      const uniqueAccountHeads = accountHeads.filter((item: any, i: number, ar: any) => ar.indexOf(item) === i);
      this.setState({
        monthlyItems: monthlyItems.value,
        monthsList: monthsList,
        accountHeads: uniqueAccountHeads,
      })

    } catch (error) {
      this.setState({ errorMessage: error.message });
    }
  }

  public componentDidMount(): void {
    this.getLast12MonthListData();
  }

  public render(): React.ReactElement<IMonthWiseExpensesProps> {
    return (
      <Box>
        {this.state.errorMessage && <Alert severity="error"> {this.state.errorMessage}</Alert>}
        {this.renderMonthWiseExpenseView()}
      </Box>
    );
  }

  private getMonthsWithStartEndDates(startDate: Date, endDate: Date): any[] {
    const months: any[] = [];
    let currentDate = new Date(startDate.toISOString());

    while (currentDate <= endDate) {
      const startOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
      const endOfMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0);

      months.push({
        start: startOfMonth.toISOString(),
        end: endOfMonth.toISOString(),
        shortName: startOfMonth.toLocaleString('default', { month: 'short' }),
        year: startOfMonth.getFullYear(),
        densibleAmount: 0,
        webTacklesAmount: 0,
        delivererAmount: 0,
      });

      // Move to the next month
      currentDate.setMonth(currentDate.getMonth() + 1);
    }

    return months;
  }

  private renderMonthWiseExpenseView(): React.ReactNode {
    const data = this.state.monthlyItems
    const accountHeads = this.state.accountHeads;

    const monthsList: any[] = this.state.monthsList
    if (data.length === 0) {
      return '';
    }
    for (let i = 0; i < monthsList.length; i++) {
      const ele: any = monthsList[i];
      const monthData = data.filter((row: any) => new Date(row.Date) > new Date(ele.start) && new Date(row.Date) < new Date(ele.end));

      for (let accInd = 0; accInd < accountHeads.length; accInd++) {
        const head = accountHeads[accInd];
        monthsList[i][head] = monthData.filter((row: any) => row.AccountHead === head)
          .map((e: any) => e.Debit).reduce((sum, current) => sum + current, 0);

      }
    }

    return (
      <div>
        <h2>Month Wise Expenses</h2>
        <table className={styles.strippedTable}>
          <thead>
            <tr>
              <th style={{ textAlign: "left" }}>Month</th>
              {
                accountHeads.map((field: string) => (
                  <th style={{ width: "76px" }} key={field}>{field}</th>
                ))
              }
            </tr>
          </thead>
          <tbody>
            {
              monthsList.map((month: any, index) => (
                <tr key={index}>
                  <td>{month.shortName} {month.year}</td>
                  {
                    accountHeads.map((field: string) => (
                      <td style={{ textAlign: "center" }} key={field}>{month[field].toFixed(2)}</td>
                    ))
                  }
                </tr>
              ))
            }
          </tbody>
        </table>
      </div>
    );
  }
}
