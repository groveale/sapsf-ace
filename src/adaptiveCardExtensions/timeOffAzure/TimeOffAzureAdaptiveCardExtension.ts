import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { LoadingCardView } from './cardView/LoadingCardView'
import { QuickView } from './quickView/QuickView';
import { TimeOffAzurePropertyPane } from './TimeOffAzurePropertyPane';
import { AadHttpClient } from '@microsoft/sp-http';
import { ITimeAccount } from './models/ITimeAccount'
import * as strings from 'TimeOffAdaptiveCardExtensionStrings';

export interface ITimeOffAzureAdaptiveCardExtensionProps {
  title: string;
}

export interface ITimeOffAzureAdaptiveCardExtensionState {
  timeOffAccounts: ITimeAccount[];
  daysAvailable: string;
  description: string;
  sapUserName: string;
  daysUntilNexTimeOff: string;
  currentlyOnLeave: Boolean
  cardState: CardState
}

enum CardState {
  Unconfigured = 1,
  Loading,
  Error,
  Loaded,
};

const CARD_VIEW_REGISTRY_ID: string = 'TimeOffAzure_CARD_VIEW';
const LOADING_CARD_VIEW_REGISTRY_ID: string = 'LOADING_CARD_VIEW';
const ERROR_CARD_VIEW_REGISTRY_ID: string = 'ERROR_CARD_VIEW';
const UNCONFIGURED_CARD_VIEW_REGISTRY_ID: string = 'UNCONFIGURED_CARD_VIEW';

export const QUICK_VIEW_REGISTRY_ID: string = 'TimeOffAzure_QUICK_VIEW';

export default class TimeOffAzureAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITimeOffAzureAdaptiveCardExtensionProps,
  ITimeOffAzureAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: TimeOffAzurePropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
        timeOffAccounts: [],
        daysAvailable: "Calculating days",
        description: strings.Description,
        sapUserName: "",
        daysUntilNexTimeOff: "",
        currentlyOnLeave: false,
        cardState: 2
     };

    this.cardNavigator.register(LOADING_CARD_VIEW_REGISTRY_ID, () => new LoadingCardView());
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve(); // this._fetchData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'TimeOffAzure-property-pane'*/
      './TimeOffAzurePropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.TimeOffAzurePropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    switch (this.state.cardState)
    {
      case CardState.Loaded:
        return CARD_VIEW_REGISTRY_ID;
      case CardState.Error:
        return ERROR_CARD_VIEW_REGISTRY_ID;
      case CardState.Loading:
        return LOADING_CARD_VIEW_REGISTRY_ID;
      case CardState.Unconfigured:
        return UNCONFIGURED_CARD_VIEW_REGISTRY_ID;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  private _fetchData(): Promise<void> {
    return this.context.aadHttpClientFactory
      .getClient('018ef16d-b0c0-45b7-b383-ab8718c63d9a')
      .then(client => client.get('https://spfx-aces101.azurewebsites.net/api/GetSFTimeAccountBalances?$sapUserNameToSearch=sfadmin&?timeAccountType=TAT_VAC_REC', AadHttpClient.configurations.v1))
      .then(response => response.json())
      .then(balances => {
        this.setState({
          timeOffAccounts: balances
        });
      });
  }
}
