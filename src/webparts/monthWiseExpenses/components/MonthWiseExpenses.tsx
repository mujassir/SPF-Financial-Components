import * as React from "react";
import type { IMonthWiseExpensesProps } from "./IMonthWiseExpensesProps";
import { Alert, Box } from "@mui/material";
import { SPHttpClient } from "@microsoft/sp-http";
import styles from "./MonthWiseExpenses.module.scss";
import Constants from "../common/constants";
import Button from '@mui/material/Button';
import jsPDF from "jspdf";

export default class MonthWiseExpenses extends React.Component<
  IMonthWiseExpensesProps,
  {
    errorMessage: string;
    getLoader: boolean;
    monthlyItems: any[];
    accountHeads: any[];
    monthsList: any[];
  }
> {
  public state;
  context: any;
  constructor(props: IMonthWiseExpensesProps) {
    super(props);

    this.state = {
      errorMessage: "",
      getLoader: false,
      monthlyItems: [],
      accountHeads: [],
      monthsList: [],
    };
  }

  private async getLast12MonthListData(): Promise<void> {
    let listUrl = "";
    let listTitle = "";

    if (this.props.listUrl) listUrl = this.props.listUrl;
    if (!listUrl || listUrl.length === 0) {
      this.setState({ errorMessage: "Please add SPFX WEB URL." });
      return;
    }

    if (this.props.lists) listTitle = this.props.lists.title;
    if (!listTitle || listTitle.length === 0) {
      this.setState({ errorMessage: "Please select a list first." });
      return;
    }

    try {
      const currentDate = new Date();
      const startDate = new Date(currentDate.toISOString());
      startDate.setMonth(currentDate.getMonth() - 11);
      startDate.setDate(1);
      const endDate = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth() + 1,
        0
      );

      const monthlyItemsEndpoint = `${listUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$filter=Date ge '${startDate.toISOString()}' and Date le '${endDate.toISOString()}'&$top=${Constants.Defaults.MaxPageSize
        }`;
      const monthlyItemsResponse = await this.props.context.spHttpClient.get(
        monthlyItemsEndpoint,
        SPHttpClient.configurations.v1
      );
      const monthlyItems = await monthlyItemsResponse.json();
      if (monthlyItems?.error) throw new Error(monthlyItems.error.message);

      const monthsList: any = this.getMonthsWithStartEndDates(
        startDate,
        endDate
      );

      const accountHeads: any = monthlyItems.value.map(
        (e: any) => e.AccountHead
      );
      const uniqueAccountHeads = accountHeads.filter(
        (item: any, i: number, ar: any) => ar.indexOf(item) === i
      );
      this.setState({
        monthlyItems: monthlyItems.value,
        monthsList: monthsList,
        accountHeads: uniqueAccountHeads,
      });
    } catch (error) {
      this.setState({ errorMessage: error.message });
    }
  }

  public async componentDidMount(): Promise<void> {
    await this.getLast12MonthListData();
  }
  public render(): React.ReactElement<IMonthWiseExpensesProps> {
    const layout: string = 'h';
    return (
      <Box>
        {this.state.errorMessage && (
          <Alert severity="error"> {this.state.errorMessage}</Alert>
        )}
        {layout === 'v' ? this.renderMonthWiseExpenseVerticalView : this.renderMonthWiseExpenseHorizentalView()}
      </Box>
    );
  }

  private getMonthsWithStartEndDates(startDate: Date, endDate: Date): any[] {
    const months: any[] = [];
    let currentDate = new Date(startDate.toISOString());

    while (currentDate <= endDate) {
      const startOfMonth = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth(),
        1,
        0, 0, 0
      );
      const endOfMonth = new Date(
        currentDate.getFullYear(),
        currentDate.getMonth() + 1,
        0,
        23, 59, 59
      );

      months.unshift({
        start: startOfMonth.toISOString(),
        end: endOfMonth.toISOString(),
        shortName: startOfMonth.toLocaleString("default", { month: "short" }),
        year: startOfMonth.getFullYear(),
      });

      // Move to the next month
      currentDate.setMonth(currentDate.getMonth() + 1);
    }

    return months;
  }

  private renderMonthWiseExpenseHorizentalView(): React.ReactNode {
    const data = this.state.monthlyItems;
    const accountHeads = this.state.accountHeads;

    const monthsList: any[] = this.state.monthsList;
    if (data.length === 0) {
      return "";
    }

    let totalAccWiseBalances: any = {
      totalDebit: 0,
      totalCredit: 0,
    }

    for (let i = 0; i < monthsList.length; i++) {
      const ele: any = monthsList[i];
      const monthData = data.filter(
        (row: any) =>
          new Date(row.Date) >= new Date(ele.start) &&
          new Date(row.Date) <= new Date(ele.end)
      );
      let totalDebit = 0;
      let totalCredit = 0;

      for (let accInd = 0; accInd < accountHeads.length; accInd++) {
        const head = accountHeads[accInd];
        const accDebit = monthData
          .filter((row: any) => row.AccountHead === head)
          .map((e: any) => e.Debit)
          .reduce((sum, current) => sum + current, 0);
        monthsList[i][head] = accDebit;

        const accCredit = monthData
          .filter((row: any) => row.AccountHead === head)
          .map((e: any) => e.Credit)
          .reduce((sum, current) => sum + current, 0);
        monthsList[i][head] = { debit: accDebit, credit: accCredit };
        totalDebit += accDebit;
        totalCredit += accCredit;
        if (!totalAccWiseBalances[head]) totalAccWiseBalances[head] = { credit: 0, debit: 0 }
        totalAccWiseBalances[head].credit += accCredit
        totalAccWiseBalances[head].debit += accDebit
        totalAccWiseBalances.totalCredit += accCredit
        totalAccWiseBalances.totalDebit += accDebit
      }
      monthsList[i].total = { debit: totalDebit, credit: totalCredit };;
    }

    return (
      <div style={{ overflow: 'auto' }}>
        <h2>
          Month Wise Expenses &nbsp; &nbsp;
          <Button className={`d-print-none`} variant="outlined" onClick={() => this.generatePDF()}>Generate PDF</Button>
        </h2>
        <table id="monthWiseExpenseTable" className={`${styles.strippedTable} ${styles.gridTable}`}>
          <thead>
            <tr>
              <th colSpan={2}>Company</th>
              {monthsList.map((month: any) => (
                <th style={{ whiteSpace: "nowrap" }}>
                  {month.shortName} {month.year}
                </th>
              ))}
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            {accountHeads.map((field: string) => (
              <React.Fragment key={field}>
                <tr>
                  <th rowSpan={3}>{field}</th>
                  <td>Credit</td>
                  {monthsList.map((month: any) => (
                    <td>
                      {month[field].credit.toFixed(2)}
                    </td>
                  ))}
                  <td>
                    {totalAccWiseBalances[field].credit.toFixed(2)}
                  </td>
                </tr>
                <tr>
                  <td>Debit</td>
                  {monthsList.map((month: any) => (
                    <td>
                      {month[field].debit.toFixed(2)}
                    </td>
                  ))}
                  <td>
                    {totalAccWiseBalances[field].debit.toFixed(2)}
                  </td>
                </tr>
                <tr>
                  <td>Balance</td>
                  {monthsList.map((month: any) => (
                    <th>
                      {(month[field].credit - month[field].debit).toFixed(2)}
                    </th>
                  ))}
                  <th>
                    {(totalAccWiseBalances[field].credit - totalAccWiseBalances[field].debit).toFixed(2)}
                  </th>
                </tr>
                <tr><td colSpan={1000}>&nbsp;</td></tr>
              </React.Fragment>
            ))}

            <tr>
              <th rowSpan={3}>Total</th>
              <td>Credit</td>
              {monthsList.map((month: any) => (
                <td>
                  {month.total.credit.toFixed(2)}
                </td>
              ))}
              <td>
                {totalAccWiseBalances.totalCredit.toFixed(2)}
              </td>
            </tr>
            <tr>
              <td>Debit</td>
              {monthsList.map((month: any) => (
                <td>
                  {month.total.debit.toFixed(2)}
                </td>
              ))}
              <td>
                {totalAccWiseBalances.totalDebit.toFixed(2)}
              </td>
            </tr>
            <tr>
              <td>Balance</td>
              {monthsList.map((month: any) => (
                <th>
                  {(month.total.credit - month.total.debit).toFixed(2)}
                </th>
              ))}
              <th>
                {(totalAccWiseBalances.totalCredit - totalAccWiseBalances.totalDebit).toFixed(2)}
              </th>
            </tr>
          </tbody>
        </table>
      </div>
    );
  }

  private renderMonthWiseExpenseVerticalView(): React.ReactNode {
    const data = this.state.monthlyItems;
    const accountHeads = this.state.accountHeads;

    const monthsList: any[] = this.state.monthsList;
    if (data.length === 0) {
      return "";
    }

    for (let i = 0; i < monthsList.length; i++) {
      const ele: any = monthsList[i];
      const monthData = data.filter(
        (row: any) =>
          new Date(row.Date) >= new Date(ele.start) &&
          new Date(row.Date) <= new Date(ele.end)
      );
      let totalDebit = 0;
      let totalCredit = 0;

      for (let accInd = 0; accInd < accountHeads.length; accInd++) {
        const head = accountHeads[accInd];
        const accDebit = monthData
          .filter((row: any) => row.AccountHead === head)
          .map((e: any) => e.Debit)
          .reduce((sum, current) => sum + current, 0);
        monthsList[i][head] = accDebit;

        const accCredit = monthData
          .filter((row: any) => row.AccountHead === head)
          .map((e: any) => e.Credit)
          .reduce((sum, current) => sum + current, 0);
        monthsList[i][head] = { debit: accDebit, credit: accCredit };
        totalDebit += accDebit;
        totalCredit += accDebit;
      }
      monthsList[i].Total = { debit: totalDebit, credit: totalCredit };;
    }

    return (
      <div>
        <h2>Month Wise Expenses</h2>
        <table className={`${styles.strippedTable} ${styles.gridTable}`}>
          <thead>
            <tr>
              <th style={{ textAlign: "left" }}>Month</th>
              {accountHeads.map((field: string) => (
                <th colSpan={3} style={{ width: "76px" }} key={field}>
                  {field}
                </th>
              ))}
              <th colSpan={3} style={{ textAlign: "center" }}>Total</th>
            </tr>
            <tr>
              <th>&nbsp;</th>
              {accountHeads.map((field: string) => (
                <React.Fragment key={field}>
                  <th>Credit</th>
                  <th>Debit</th>
                  <th>Balance</th>
                </React.Fragment>
              ))}
              <th>Credit</th>
              <th>Debit</th>
              <th>Balance</th>
            </tr>
          </thead>
          <tbody>
            {monthsList.map((month: any, index) => (
              <tr key={index}>
                <td>
                  {month.shortName} {month.year}
                </td>
                {accountHeads.map((field: string) => (
                  <React.Fragment key={field}>
                    <td style={{ textAlign: "center" }} key={field}>
                      {month[field].credit.toFixed(2)}
                    </td>
                    <td style={{ textAlign: "center" }} key={field}>
                      {month[field].debit.toFixed(2)}
                    </td>
                    <td style={{ textAlign: "center" }} key={field}>
                      {(month[field].credit - month[field].debit).toFixed(2)}
                    </td>
                  </React.Fragment>
                ))}
                <td style={{ textAlign: "center" }}>
                  {month.Total.credit.toFixed(2)}
                </td>
                <td style={{ textAlign: "center" }}>
                  {month.Total.debit.toFixed(2)}
                </td>
                <td style={{ textAlign: "center" }}>
                  {(month.Total.credit - month.Total.debit).toFixed(2)}
                </td>
              </tr>
            ))}
            {/* <tr>
              <td>Total</td>
              {accountHeads.map((field: string) => (
                <td style={{ textAlign: "center" }} key={field}>
                  {monthsList
                    .reduce((sum, month) => sum + month[field], 0)
                    .toFixed(2)}
                </td>
              ))}
              <td style={{ textAlign: "center" }}>
                {monthsList
                  .reduce((sum, month) => sum + month.Total, 0)
                  .toFixed(2)}
              </td>
            </tr> */}
          </tbody>
        </table>
      </div>
    );
  }

  private generatePDF = async () => {
    try {
      let htmlTableElement = document.getElementById('monthWiseExpenseTable');
      if (!htmlTableElement) return

      const contentWidth = htmlTableElement.scrollWidth;
      const contentHeight = htmlTableElement.scrollHeight;
      const doc = new jsPDF({
        orientation: 'landscape',
        unit: 'px',
        format: [contentWidth + 40, contentHeight + 40]
      });
      doc.html(htmlTableElement, {
        callback: function (doc) {
          doc.save(`Month Wise Expenses - ${(new Date()).toDateString()}.pdf`);
        },
        x: 20,
        y: 20,
        html2canvas: {
          scale: 1,
          width: contentWidth,
          windowWidth: contentWidth
        },
        autoPaging: 'text'
      });

    } catch (error) {
      console.error('Error generating PDF:', error);
    }
  };
}
