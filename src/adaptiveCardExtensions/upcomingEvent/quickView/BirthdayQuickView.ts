import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'UpcomingEventAdaptiveCardExtensionStrings';
import {
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState
} from '../IUpcomingEventAdaptiveCardExtensionProps';
import UpcomingEventAdaptiveCardExtension from '../UpcomingEventAdaptiveCardExtension';
import { IEventOccurrence } from '../../../shared/services';

// Default birthday image - using relative path from quickView to webparts/holidaysBirthdays/assets
const defaultBirthdayImage = require('../../../webparts/holidaysBirthdays/assets/default-birthday.jpg');

export interface IBirthdayQuickViewData {
  title: string;
  birthdays: Array<{
    title: string;
    date: string;
    daysUntil: string;
    imageUrl: string;
  }>;
  showBranding: boolean;
  companyName: string;
  isLoading: boolean;
  errorMessage: string | null;
  noBirthdaysMessage: string;
  hasBirthdays: boolean;
}

export class BirthdayQuickView extends BaseAdaptiveCardQuickView<
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState,
  IBirthdayQuickViewData
> {
  private _birthdays: IEventOccurrence[] = [];
  private _isLoading: boolean = true;
  private _error: string | null = null;

  public async onInit(): Promise<void> {
    await this._loadBirthdays();
  }

  private async _loadBirthdays(): Promise<void> {
    const ace = this.context as unknown as UpcomingEventAdaptiveCardExtension;
    const dataService = ace.getDataService();

    if (!dataService) {
      this._error = 'Data service not available';
      this._isLoading = false;
      return;
    }

    try {
      this._isLoading = true;
      const count = this.properties.birthdayCount || 5;
      this._birthdays = await dataService.getNextBirthdays(count);
      this._isLoading = false;
      this._error = null;
    } catch (error) {
      console.error('Error loading birthdays:', error);
      this._error = error instanceof Error ? error.message : 'Failed to load birthdays';
      this._isLoading = false;
    }
  }

  public get data(): IBirthdayQuickViewData {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const birthdays = this._birthdays.map(b => {
      const daysUntil = this._calculateDaysUntil(b.date);
      let daysUntilText = '';

      if (daysUntil === 0) {
        daysUntilText = 'Today! 🎉';
      } else if (daysUntil === 1) {
        daysUntilText = 'Tomorrow';
      } else if (daysUntil <= 7) {
        daysUntilText = `in ${daysUntil} days`;
      } else if (daysUntil <= 14) {
        daysUntilText = 'in about 2 weeks';
      } else if (daysUntil <= 30) {
        daysUntilText = `in ${Math.ceil(daysUntil / 7)} weeks`;
      } else {
        daysUntilText = `in ${Math.ceil(daysUntil / 30)} months`;
      }

      return {
        title: b.title,
        date: b.date.toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'long',
          day: 'numeric'
        }),
        daysUntil: daysUntilText,
        imageUrl: b.imageUrl || defaultBirthdayImage
      };
    });

    return {
      title: strings.BirthdaysTitle,
      birthdays: birthdays,
      showBranding: this.properties.showBranding !== false,
      companyName: strings.CompanyName,
      isLoading: this._isLoading,
      errorMessage: this._error,
      noBirthdaysMessage: strings.NoEventsMessage,
      hasBirthdays: birthdays.length > 0
    };
  }

  private _calculateDaysUntil(eventDate: Date): number {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const eventDateOnly = new Date(eventDate);
    eventDateOnly.setHours(0, 0, 0, 0);
    const diffTime = eventDateOnly.getTime() - today.getTime();
    return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  }

  public get template(): ISPFxAdaptiveCard {
    return {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.5',
      body: [
        // Header
        {
          type: 'Container',
          items: [
            {
              type: 'TextBlock',
              text: '🎂 ${title}',
              size: 'Large',
              weight: 'Bolder',
              wrap: true
            }
          ]
        },
        // Loading indicator
        {
          type: 'TextBlock',
          text: 'Loading birthdays...',
          isSubtle: true,
          horizontalAlignment: 'Center',
          $when: '${isLoading}'
        },
        // Error message
        {
          type: 'TextBlock',
          text: '${errorMessage}',
          color: 'Attention',
          wrap: true,
          $when: '${errorMessage != null}'
        },
        // No birthdays message
        {
          type: 'TextBlock',
          text: '${noBirthdaysMessage}',
          isSubtle: true,
          horizontalAlignment: 'Center',
          $when: '${!isLoading && errorMessage == null && !hasBirthdays}'
        },
        // Birthday list
        {
          type: 'Container',
          items: [
            {
              type: 'Container',
              $data: '${birthdays}',
              separator: true,
              spacing: 'Medium',
              items: [
                {
                  type: 'ColumnSet',
                  columns: [
                    {
                      type: 'Column',
                      width: '48px',
                      items: [
                        {
                          type: 'Image',
                          url: '${imageUrl}',
                          size: 'Small',
                          style: 'Person'
                        }
                      ]
                    },
                    {
                      type: 'Column',
                      width: 'stretch',
                      items: [
                        {
                          type: 'TextBlock',
                          text: '${title}',
                          weight: 'Bolder',
                          wrap: true
                        },
                        {
                          type: 'TextBlock',
                          text: '${date}',
                          isSubtle: true,
                          spacing: 'None',
                          wrap: true
                        },
                        {
                          type: 'TextBlock',
                          text: '${daysUntil}',
                          isSubtle: true,
                          spacing: 'None',
                          size: 'Small',
                          color: 'Accent'
                        }
                      ],
                      verticalContentAlignment: 'Center'
                    }
                  ]
                }
              ]
            }
          ],
          $when: '${!isLoading && errorMessage == null && hasBirthdays}'
        },
        // Branding footer
        {
          type: 'Container',
          separator: true,
          spacing: 'Medium',
          items: [
            {
              type: 'TextBlock',
              text: '${companyName}',
              isSubtle: true,
              size: 'Small',
              horizontalAlignment: 'Center'
            }
          ],
          $when: '${showBranding}'
        }
      ]
    };
  }
}
