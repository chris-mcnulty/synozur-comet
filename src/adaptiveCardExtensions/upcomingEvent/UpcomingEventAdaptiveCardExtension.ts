import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";

import { CardView } from './cardView/CardView';
import { HolidayQuickView } from './quickView/HolidayQuickView';
import { BirthdayQuickView } from './quickView/BirthdayQuickView';
import { IUpcomingEventAdaptiveCardExtensionProps, IUpcomingEventAdaptiveCardExtensionState } from './IUpcomingEventAdaptiveCardExtensionProps';
import { DataService, IEventOccurrence } from '../../shared/services';
import * as strings from 'UpcomingEventAdaptiveCardExtensionStrings';

// Default images - using relative path from adaptiveCardExtensions/upcomingEvent to webparts/holidaysBirthdays/assets
const defaultHolidayImage = require('../../webparts/holidaysBirthdays/assets/default-holiday.jpg');
const defaultBirthdayImage = require('../../webparts/holidaysBirthdays/assets/default-birthday.jpg');
const defaultChristmasImage = require('../../webparts/holidaysBirthdays/assets/default-christmas.jpg');
const defaultIndependenceImage = require('../../webparts/holidaysBirthdays/assets/default-independence.jpg');

export const CARD_VIEW_REGISTRY_ID: string = 'UpcomingEvent_CARD_VIEW';
export const HOLIDAY_QUICK_VIEW_REGISTRY_ID: string = 'UpcomingEvent_HOLIDAY_QUICK_VIEW';
export const BIRTHDAY_QUICK_VIEW_REGISTRY_ID: string = 'UpcomingEvent_BIRTHDAY_QUICK_VIEW';

export default class UpcomingEventAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState
> {
  private _dataService: DataService | null = null;

  public onInit(): Promise<void> {
    // Initialize state
    this.state = {
      nextEventTitle: '',
      nextEventDate: null,
      nextEventImageUrl: '',
      daysUntil: 0,
      isLoading: true,
      error: null,
      selectedYear: new Date().getFullYear()
    };

    // Register card views
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(HOLIDAY_QUICK_VIEW_REGISTRY_ID, () => new HolidayQuickView());
    this.quickViewNavigator.register(BIRTHDAY_QUICK_VIEW_REGISTRY_ID, () => new BirthdayQuickView());

    // Initialize PnP
    sp.setup({
      spfxContext: this.context as any
    });

    // Initialize data service and load data
    this._dataService = new DataService(this.properties.listTitle || 'HolidaysAndBirthdays');
    
    return this._loadData();
  }

  public get iconProperty(): string {
    return this.properties.displayMode === 'Birthday' ? '🎂' : '🎄';
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return Promise.resolve();
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isBirthdayMode = this.properties.displayMode === 'Birthday';

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
                PropertyPaneDropdown('displayMode', {
                  label: strings.DisplayModeFieldLabel,
                  options: [
                    { key: 'Holiday', text: strings.DisplayModeHoliday },
                    { key: 'Birthday', text: strings.DisplayModeBirthday }
                  ],
                  selectedKey: this.properties.displayMode || 'Holiday'
                }),
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel,
                  value: this.properties.listTitle || 'HolidaysAndBirthdays'
                }),
                PropertyPaneToggle('showDaysUntil', {
                  label: strings.ShowDaysUntilFieldLabel,
                  checked: this.properties.showDaysUntil !== false
                }),
                PropertyPaneToggle('showBranding', {
                  label: strings.ShowBrandingFieldLabel,
                  checked: this.properties.showBranding !== false
                })
              ]
            },
            // Birthday-specific settings (only show when in Birthday mode)
            ...(isBirthdayMode ? [{
              groupName: strings.BirthdaySettingsGroupName,
              groupFields: [
                PropertyPaneSlider('birthdayCount', {
                  label: strings.BirthdayCountFieldLabel,
                  min: 3,
                  max: 10,
                  step: 1,
                  value: this.properties.birthdayCount || 5
                })
              ]
            }] : [])
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listTitle' || propertyPath === 'displayMode') {
      // Reinitialize data service if list title changes
      this._dataService = new DataService(this.properties.listTitle || 'HolidaysAndBirthdays');
      this._loadData();
    }
  }

  /**
   * Load the next event data from the SharePoint list
   */
  private async _loadData(): Promise<void> {
    if (!this._dataService) {
      return;
    }

    try {
      this.setState({
        isLoading: true,
        error: null
      });

      const eventType = this.properties.displayMode === 'Birthday' ? 'Birthday' : 'Holiday';
      const nextEvent = await this._dataService.getNextEvent(eventType);

      if (nextEvent) {
        const daysUntil = this._calculateDaysUntil(nextEvent.date);
        const imageUrl = this._getImageUrl(nextEvent);

        this.setState({
          nextEventTitle: nextEvent.title,
          nextEventDate: nextEvent.date,
          nextEventImageUrl: imageUrl,
          daysUntil: daysUntil,
          isLoading: false,
          error: null
        });
      } else {
        this.setState({
          nextEventTitle: strings.NoEventsMessage,
          nextEventDate: null,
          nextEventImageUrl: this._getDefaultImage(),
          daysUntil: 0,
          isLoading: false,
          error: null
        });
      }
    } catch (error) {
      console.error('Error loading ACE data:', error);
      this.setState({
        nextEventTitle: strings.ErrorMessage,
        nextEventDate: null,
        nextEventImageUrl: this._getDefaultImage(),
        daysUntil: 0,
        isLoading: false,
        error: error instanceof Error ? error.message : 'Unknown error'
      });
    }
  }

  /**
   * Calculate days until the event
   */
  private _calculateDaysUntil(eventDate: Date): number {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const eventDateOnly = new Date(eventDate);
    eventDateOnly.setHours(0, 0, 0, 0);
    const diffTime = eventDateOnly.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  }

  /**
   * Get the image URL for an event, with fallback to defaults
   */
  private _getImageUrl(event: IEventOccurrence): string {
    if (event.imageUrl) {
      return event.imageUrl;
    }
    return this._getDefaultImageForEvent(event);
  }

  /**
   * Get the default image for an event based on type and date
   */
  private _getDefaultImageForEvent(event: IEventOccurrence): string {
    if (event.eventType === 'Holiday') {
      const month = event.date.getMonth();
      const day = event.date.getDate();

      // Christmas (December 25)
      if (month === 11 && day === 25) {
        return defaultChristmasImage;
      }

      // Independence Day (July 4)
      if (month === 6 && day === 4) {
        return defaultIndependenceImage;
      }

      return defaultHolidayImage;
    }

    return defaultBirthdayImage;
  }

  /**
   * Get the default image based on display mode
   */
  private _getDefaultImage(): string {
    return this.properties.displayMode === 'Birthday' ? defaultBirthdayImage : defaultHolidayImage;
  }

  /**
   * Get the DataService instance (used by Quick Views)
   */
  public getDataService(): DataService | null {
    return this._dataService;
  }

  /**
   * Update the selected year for holiday navigation
   */
  public setSelectedYear(year: number): void {
    this.setState({
      selectedYear: year
    });
  }
}
