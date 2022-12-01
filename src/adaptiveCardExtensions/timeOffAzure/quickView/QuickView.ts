import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TimeOffAzureAdaptiveCardExtensionStrings';
import { ITimeAccount } from '../models/ITimeAccount';
import { HISTORY_QUICK_VIEW_REGISTRY_ID, ITimeOffAzureAdaptiveCardExtensionProps, ITimeOffAzureAdaptiveCardExtensionState } from '../TimeOffAzureAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
  items: ITimeAccount[]
  faqsLink: string
}

export class QuickView extends BaseAdaptiveCardView<
  ITimeOffAzureAdaptiveCardExtensionProps,
  ITimeOffAzureAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      items: this.state.timeOffAccounts,
      faqsLink: this.properties.FAQLink
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/BalancesTimeOff.json');
  }

  public onAction(action: IActionArguments): void {
    if (action.type === 'Submit') {
      const { id, newIndex } = action.data;
      if (id === 'viewHistory') {
        // false is important if not updating the state
        this.quickViewNavigator.push(HISTORY_QUICK_VIEW_REGISTRY_ID, false)
      }
    }
  }
}