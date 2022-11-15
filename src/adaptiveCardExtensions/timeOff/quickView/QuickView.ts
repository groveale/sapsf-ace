import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TimeOffAdaptiveCardExtensionStrings';
import { ITimeAccount } from '../models/ITimeAccount';
import { ITimeOffAdaptiveCardExtensionProps, ITimeOffAdaptiveCardExtensionState } from '../TimeOffAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
  items: ITimeAccount[];
  faqsLink: string;
}

export class QuickView extends BaseAdaptiveCardView<
  ITimeOffAdaptiveCardExtensionProps,
  ITimeOffAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      items: this.state.timeOffAccounts,
      faqsLink: "https://google.com"
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/BalancesTemplate.json');
  }
}