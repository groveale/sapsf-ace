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
                  label: strings.ListFieldLabel
                }),
                PropertyPaneTextField('SAPAdProperty', {
                  label: strings.SAPAdFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
