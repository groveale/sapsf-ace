import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'TimeOffAdaptiveCardExtensionStrings';

export class TimeOffPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: strings.ListFieldLabel,
                  value: strings.ListTitle
                }),
                PropertyPaneTextField('SAPAdField', {
                  label: strings.SAPAdFieldLabel,
                  value: strings.SAPAdField
                })
              ]
            },
            {
              groupFields: [
                PropertyPaneTextField('SAPSFHostname', {
                  label: strings.SAPSFHostnameLabel,
                  value: strings.SAPSFHostname
                }),
                PropertyPaneTextField('SAPSFAPIKey', {
                  label: strings.SAPSFAPIKeyLabel,
                  value: strings.SAPSFAPIKey
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
