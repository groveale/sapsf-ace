import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { TimeOffPropertyPane } from './TimeOffPropertyPane';
import { SPHttpClient, HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { ITimeAccount, TimeAccount } from './models/ITimeAccount';
import * as strings from 'TimeOffAdaptiveCardExtensionStrings';

export interface ITimeOffAdaptiveCardExtensionProps {
  title: string;
  SAPSFHostname: string;
  SAPSFAPIKey: string;
  listTitle: string;
  SAPAdField: string;
  FAQLink: string
}

export interface ITimeOffAdaptiveCardExtensionState {
  timeOffAccounts: ITimeAccount[];
  daysAvailable: string;
  description: string;
  sapUserName: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'TimeOff_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'TimeOff_QUICK_VIEW';

export default class TimeOffAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ITimeOffAdaptiveCardExtensionProps,
  ITimeOffAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: TimeOffPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
      timeOffAccounts: [],
      daysAvailable: "Calculating days",
      description: strings.Description,
      sapUserName: ""
    };

    await this.getSAPSFUserNameFromAAD()
    await this.getTimeAccountsFromSPOList()
    let id = this.state.sapUserName
    await this.getTimeAccountDetailsFromSAPSF()
    
    // Register Card Views
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'TimeOff-property-pane'*/
      './TimeOffPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.TimeOffPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
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
    .then((response: any): TimeAccount[] => {
      let timeAccountSPOArray: TimeAccount[] = []
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
            const timeAccount = new TimeAccount(item['ID'], item['Title'], item['HolidayTypeDescription'], item['HolidayTypeSAPIdentifier'], item['HolidayTimeTypeSAPIdentifier'], iconImage, "", 0, 0);
            timeAccountSPOArray.push(timeAccount)
          })
        }
      }
      catch
      {
        console.log("Error")
        this.setState(
        { 
          daysAvailable: "SPO Error",
          description: "Check Time Account list configuration"
        }
        );
      }
      
      return timeAccountSPOArray
      })
      .then((items: TimeAccount[]) => this.setState(
        { timeOffAccounts: items }
      ));
  }


  private async getSAPSFUserNameFromAAD() : Promise<void>  {
    return this.context.msGraphClientFactory
      .getClient('3')
      .then(client => client.api('me').select(`${this.properties.SAPAdField}`).get())
      .then((sapSAPUserFromAAD: any) => {
        this.setState({
          sapUserName: sapSAPUserFromAAD[`${this.properties.SAPAdField}`]
      });
    })
  }

  private async getTimeAccountDetailsFromSAPSF() : Promise<void>  {
    let totalBalance = 0
    let today = new Date()
    console.log(this.state.timeOffAccounts.length);
    this.state.timeOffAccounts.forEach(timeAccount => {
      this.context.httpClient
      .get(`${this.properties.SAPSFHostname}/odata/v2/EmpTimeAccountBalance?$filter=userId eq '${this.state.sapUserName}' and timeAccountType eq '${timeAccount.sapIdentifierTAT}'&$format=json`, HttpClient.configurations.v1,
        {
          headers: [
            ['accept', 'application/json'],
            ['APIKey', this.properties.SAPSFAPIKey]
          ]
        })
      .then((res: HttpClientResponse): Promise<any> => {
        return res.json();
      })
      .then((response: any): void => {
        console.log(response.d.results.length);
        
        let balancesFromSAPSF: [] = response.d.results
        let accountBalance = 0
        if (balancesFromSAPSF)
        {
          balancesFromSAPSF.forEach(balance => {
            accountBalance += +balance['balance']
          });
        }
        timeAccount.balanceDaysString = accountBalance.toString()
        totalBalance += accountBalance
        })
        .then(() => this.setState(
          { daysAvailable: totalBalance.toString() + " days" }
        ))
      //   .then(() => this.context.httpClient
      //   .get(`${this.properties.SAPSFHostname}/odata/v2/EmployeeTime?$filter=userId eq '${this.state.sapUserName}' and timeType eq '${timeAccount.sapIdentifierTT} endDate gt datetime'${today}'&$select=approvalStatus,quantityInHours,quantityInDays,startDate,endDate,timeType&$format=json`, HttpClient.configurations.v1,
      //     {
      //       headers: [
      //         ['accept', 'application/json'],
      //         ['APIKey', this.properties.SAPSFAPIKey]
      //       ]
      //     })
      //   .then((res: HttpClientResponse): Promise<any> => {
      //     return res.json();
      //   })
      //   .then((response: any): void => {
      //     console.log(response.d.results.length);
      //   })
      // );

    })
  }
}
