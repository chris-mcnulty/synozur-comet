import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HolidaysBirthdaysWebPartStrings';
import HolidaysBirthdays from './components/HolidaysBirthdays';
import { IHolidaysBirthdaysWebPartProps } from './IHolidaysBirthdaysWebPartProps';

export default class HolidaysBirthdaysWebPart extends BaseClientSideWebPart<IHolidaysBirthdaysWebPartProps> {

  public render(): void {
    const element: React.ReactElement = React.createElement(
      HolidaysBirthdays,
      {
        context: this.context,
        daysDefault: this.properties.daysDefault ?? 30,
        daysExpanded: this.properties.daysExpanded ?? 365,
        listTitle: this.properties.listTitle ?? 'HolidaysAndBirthdays',
        showImages: this.properties.showImages !== false,
        showTypeBadges: this.properties.showTypeBadges !== false,
        allowListProvisioning: this.properties.allowListProvisioning !== false
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneSlider('daysDefault', {
                  label: strings.DaysDefaultFieldLabel,
                  min: 7,
                  max: 90,
                  step: 1,
                  value: this.properties.daysDefault ?? 30
                }),
                PropertyPaneSlider('daysExpanded', {
                  label: strings.DaysExpandedFieldLabel,
                  min: 60,
                  max: 730,
                  step: 1,
                  value: this.properties.daysExpanded ?? 365
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel,
                  value: this.properties.listTitle ?? 'HolidaysAndBirthdays'
                }),
                PropertyPaneToggle('showImages', {
                  label: strings.ShowImagesFieldLabel,
                  checked: this.properties.showImages !== false
                }),
                PropertyPaneToggle('showTypeBadges', {
                  label: strings.ShowTypeBadgesFieldLabel,
                  checked: this.properties.showTypeBadges !== false
                }),
                PropertyPaneToggle('allowListProvisioning', {
                  label: strings.AllowListProvisioningFieldLabel,
                  checked: this.properties.allowListProvisioning !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

