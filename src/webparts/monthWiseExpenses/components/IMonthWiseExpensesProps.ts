import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMonthWiseExpensesProps {
  context: WebPartContext;
  listUrl: string;
  lists: any;
}
