import * as React from 'react';
import type { IExpenseSummaryProps } from './IExpenseSummaryProps';

export default class ExpenseSummary extends React.Component<IExpenseSummaryProps, {}> {
  
  public render(): React.ReactElement<IExpenseSummaryProps> {
    return (
      <div>Expense Summary</div>
    );
  }

}
