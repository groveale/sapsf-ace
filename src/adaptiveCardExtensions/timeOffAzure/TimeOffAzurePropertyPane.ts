import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'TimeOffAzureAdaptiveCardExtensionStrings';

export class TimeOffAzurePropertyPane {
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
          ]
        }
      ]
    };
  }
}
