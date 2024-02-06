import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGroupByField } from "../models/IGroupByField";

export interface IDailyExpenseDetailProps {
  context: WebPartContext;
  listUrl: string;
  lists: any;
  listColumns: any[];
  orderedListColumns: any[];
  groupByFields: IGroupByField[];
}
