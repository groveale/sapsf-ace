import {
    BasePrimaryTextCardView,
    IPrimaryTextCardParameters,
    IExternalLinkCardAction,
    IQuickViewCardAction,
    ICardButton
  } from '@microsoft/sp-adaptive-card-extension-base';
  import * as strings from 'TimeOffAzureAdaptiveCardExtensionStrings';
  import { ITimeOffAzureAdaptiveCardExtensionProps, ITimeOffAzureAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../TimeOffAzureAdaptiveCardExtension';
  
  export class ErrorCardView extends BasePrimaryTextCardView<ITimeOffAzureAdaptiveCardExtensionProps, ITimeOffAzureAdaptiveCardExtensionState> {
  
    public get data(): IPrimaryTextCardParameters {
      return {
        primaryText: strings.ErrorMessage,
        title: this.properties.title,
        description: this.state.exceptionMessage
      };
    }
  }
  