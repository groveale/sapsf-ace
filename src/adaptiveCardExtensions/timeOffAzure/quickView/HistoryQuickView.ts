import { BaseAdaptiveCardView, IActionArguments, ISPFxAdaptiveCard } from '@microsoft/sp-adaptive-card-extension-base';
import { ITimeOffAzureAdaptiveCardExtensionProps, ITimeOffAzureAdaptiveCardExtensionState } from '../TimeOffAzureAdaptiveCardExtension';
import { ITimeBooked } from '../models/ITimeBooked';

export interface IHistoryQuickViewData {
  pastTime: ITimeBooked[];
}

export class HistoryQuickView extends BaseAdaptiveCardView<
ITimeOffAzureAdaptiveCardExtensionProps,
ITimeOffAzureAdaptiveCardExtensionState,
IHistoryQuickViewData
> {
  public get data(): IHistoryQuickViewData {
    return {
        pastTime: this.state.timeOffAccounts[0].timeBookedPast
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/History.json');
  }

  public onAction(action: IActionArguments): void {
    if (action.type === 'Submit') {
      const { id } = action.data;
      if (id === 'back') {
        this.quickViewNavigator.pop();
      }
    }
  }
}