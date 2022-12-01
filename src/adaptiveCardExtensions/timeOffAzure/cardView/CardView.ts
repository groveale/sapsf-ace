import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TimeOffAzureAdaptiveCardExtensionStrings';
import { ITimeOffAzureAdaptiveCardExtensionProps, ITimeOffAzureAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../TimeOffAzureAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<ITimeOffAzureAdaptiveCardExtensionProps, ITimeOffAzureAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IPrimaryTextCardParameters {
    let nextTimeOff: string = "Next time off in"
    if(this.state.daysUntilNexTimeOff <= 0) {
      nextTimeOff = "Currently on leave 😎"
    }
    else if(this.state.daysUntilNexTimeOff == 10000) {
      // No leave planned
      nextTimeOff = "No time off scheduled, book it in"
    }
    else {
      nextTimeOff = `${nextTimeOff} ${this.state.daysUntilNexTimeOff} days`
    }
    return {
      primaryText: `${this.state.daysAvailable.toString()} days avalible`,
      title: this.properties.title,
      description: nextTimeOff
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
