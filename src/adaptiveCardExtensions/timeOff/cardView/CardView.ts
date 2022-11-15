import {
  BaseBasicCardView,
  BasePrimaryTextCardView,
  IBasicCardParameters,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'TimeOffAdaptiveCardExtensionStrings';
import { ITimeOffAdaptiveCardExtensionProps, ITimeOffAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../TimeOffAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<ITimeOffAdaptiveCardExtensionProps, ITimeOffAdaptiveCardExtensionState> {
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
    return {
      primaryText: this.state.daysAvailable,
      title: this.properties.title,
      description: strings.Description
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
