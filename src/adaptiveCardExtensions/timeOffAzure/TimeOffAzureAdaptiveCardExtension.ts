import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { LoadingCardView } from './cardView/LoadingCardView'
import { QuickView } from './quickView/QuickView';
import { TimeOffAzurePropertyPane } from './TimeOffAzurePropertyPane';
import { AadHttpClient } from '@microsoft/sp-http';
import { ITimeAccount } from './models/ITimeAccount'
import * as strings from 'TimeOffAdaptiveCardExtensionStrings';
import { ErrorCardView } from './cardView/ErrorCardView';
import { UnconfiguredCardView } from './cardView/UnconfiguredCardView';
import { SPHttpClient, HttpClient, HttpClientResponse } from '@microsoft/sp-http';

export interface ITimeOffAzureAdaptiveCardExtensionProps {
  title: string;
  listTitle: string;
  SAPAdField: string;
  FAQLink: string
}

export interface ITimeOffAzureAdaptiveCardExtensionState {
  timeOffAccounts: ITimeAccount[];
  daysAvailable: string;
  description: string;
  sapUserName: string;
  daysUntilNexTimeOff: string;
  currentlyOnLeave: Boolean
  cardState: CardState
  exceptionMessage: string
  configMessage: string
  loadingLog: string
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

  public async onInit(): Promise<void> {
    this.state = {
        timeOffAccounts: [],
        daysAvailable: "Calculating days",
        description: strings.Description,
        sapUserName: "",
        daysUntilNexTimeOff: "",
        currentlyOnLeave: false,
        cardState: 2,
        exceptionMessage: "try catch this",
        configMessage: "missing library",
        loadingLog: "Warming up"
     };

    this.cardNavigator.register(UNCONFIGURED_CARD_VIEW_REGISTRY_ID, () => new UnconfiguredCardView());
    this.cardNavigator.register(ERROR_CARD_VIEW_REGISTRY_ID, () => new ErrorCardView());
    this.cardNavigator.register(LOADING_CARD_VIEW_REGISTRY_ID, () => new LoadingCardView());
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    // First Step is to get SAP User Name
    await this.getSAPSFUserNameFromAAD()

    // Now Get SAP timeaccount details from SPO list
    await this.getTimeAccountsFromSPOList()

    // Call the Enterprise API


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

  private async getSAPSFUserNameFromAAD() : Promise<void>  {
    try {
        return this.context.msGraphClientFactory
        .getClient('3')
        .then(client => client.api('me').select(`${this.properties.SAPAdField}`).get())
        .then((sapSAPUserFromAAD: any) => {
          this.setState({
            sapUserName: sapSAPUserFromAAD[`${this.properties.SAPAdField}`],
            loadingLog: "Obtained SAP id: " + sapSAPUserFromAAD[`${this.properties.SAPAdField}`]
        });
      })
    }
    catch {
      this.setState(
        { 
          cardState: CardState.Error,
          exceptionMessage: "Unable to read SAP username from aad property"
        }
      );
    }
    finally {
      return Promise.resolve()
    }
  }

  private async getTimeAccountsFromSPOList() : Promise<void>  {

    return await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getByTitle('${this.properties.listTitle}')/items?$filter=(ShowInCard eq 1)&$select=HolidayTypeSAPIdentifier,HolidayTypeDescription,HolidayTypeIcon,Title,ShowInCard`,
      SPHttpClient.configurations.v1
    )
    .then((res: HttpClientResponse): Promise<any> => {
      return res.json();
    })
    .then((response: any): ITimeAccount[] => {
      let timeAccountSPOArray: ITimeAccount[] = []
      try {
        console.log(response.value.length);
        let timeAccountsFromSPO: [] = response.value
        
        if (timeAccountsFromSPO)
        {
          timeAccountsFromSPO.forEach(item => {
            // extract header image URL
            let iconImage:string = "";
  
            if (item['HolidayTypeIcon']) 
            {
              var iconJson = JSON.parse(item['HolidayTypeIcon']);
              iconImage = `${iconJson.serverUrl}${iconJson.serverRelativeUrl}`
            }

            // Create Time Account obj
            const timeAccount = {} as ITimeAccount;
            timeAccount.id = item['ID']
            timeAccount.title = item['Title']
            timeAccount.description = item['HolidayTypeDescription']
            timeAccount.sapIdentifierTAT = item['HolidayTypeSAPIdentifier'],
            timeAccount.sapIdentifierTT =  item['HolidayTimeTypeSAPIdentifier']
            timeAccount.picture = iconImage

            timeAccountSPOArray.push(timeAccount)
          })
        }
      }
      catch
      {
        console.log("Error")
        this.setState(
          { 
            cardState: CardState.Error,
            exceptionMessage: "Unable to read SharePoint list"
          }
        );
      }
      
      return timeAccountSPOArray
      })
    .then((items: ITimeAccount[]) => this.setState(
      { 
        timeOffAccounts: items,
        loadingLog: `${items.length} SAP TAT obtained from SPO list`  
      }
    ));
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
