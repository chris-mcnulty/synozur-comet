import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, IActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'UpcomingEventAdaptiveCardExtensionStrings';
import {
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState
} from '../IUpcomingEventAdaptiveCardExtensionProps';
import UpcomingEventAdaptiveCardExtension from '../UpcomingEventAdaptiveCardExtension';
import { IEventOccurrence } from '../../../shared/services';

// Default images for fallback - using relative path from quickView to webparts/holidaysBirthdays/assets
const defaultHolidayImage = require('../../../webparts/holidaysBirthdays/assets/default-holiday.jpg');
const defaultChristmasImage = require('../../../webparts/holidaysBirthdays/assets/default-christmas.jpg');
const defaultIndependenceImage = require('../../../webparts/holidaysBirthdays/assets/default-independence.jpg');

export interface IHolidayQuickViewData {
  title: string;
  selectedYear: number;
  holidays: Array<{
    title: string;
    date: string;
    imageUrl: string;
    isPast: boolean;
  }>;
  showBranding: boolean;
  companyName: string;
  hasPreviousYear: boolean;
  hasNextYear: boolean;
  previousYearLabel: string;
  nextYearLabel: string;
  isLoading: boolean;
  errorMessage: string | null;
}

export class HolidayQuickView extends BaseAdaptiveCardQuickView<
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState,
  IHolidayQuickViewData
> {
  private _holidays: IEventOccurrence[] = [];
  private _isLoading: boolean = true;
  private _error: string | null = null;

  public async onInit(): Promise<void> {
    await this._loadHolidays();
  }

  private async _loadHolidays(): Promise<void> {
    const ace = this.context as unknown as UpcomingEventAdaptiveCardExtension;
    const dataService = ace.getDataService();

    if (!dataService) {
      this._error = 'Data service not available';
      this._isLoading = false;
      return;
    }

    try {
      this._isLoading = true;
      this._holidays = await dataService.getHolidaysForYear(this.state.selectedYear);
      this._isLoading = false;
      this._error = null;
    } catch (error) {
      console.error('Error loading holidays:', error);
      this._error = error instanceof Error ? error.message : 'Failed to load holidays';
      this._isLoading = false;
    }
  }

  public get data(): IHolidayQuickViewData {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const holidays = this._holidays.map(h => ({
      title: h.title,
      date: h.date.toLocaleDateString('en-US', {
        weekday: 'long',
        month: 'long',
        day: 'numeric'
      }),
      imageUrl: this._getImageUrl(h),
      isPast: h.date < today
    }));

    return {
      title: strings.HolidaysTitle,
      selectedYear: this.state.selectedYear,
      holidays: holidays,
      showBranding: this.properties.showBranding !== false,
      companyName: strings.CompanyName,
      hasPreviousYear: true,
      hasNextYear: true,
      previousYearLabel: `◀ ${this.state.selectedYear - 1}`,
      nextYearLabel: `${this.state.selectedYear + 1} ▶`,
      isLoading: this._isLoading,
      errorMessage: this._error
    };
  }

  private _getImageUrl(event: IEventOccurrence): string {
    if (event.imageUrl) {
      return event.imageUrl;
    }

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

  public get template(): ISPFxAdaptiveCard {
    return {
      $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
      type: 'AdaptiveCard',
      version: '1.5',
      body: [
        // Header with title and year
        {
          type: 'Container',
          items: [
            {
              type: 'TextBlock',
              text: '🎄 ${title}',
              size: 'Large',
              weight: 'Bolder',
              wrap: true
            }
          ]
        },
        // Year navigation
        {
          type: 'ColumnSet',
          columns: [
            {
              type: 'Column',
              width: 'auto',
              items: [
                {
                  type: 'ActionSet',
                  actions: [
                    {
                      type: 'Action.Submit',
                      title: '${previousYearLabel}',
                      data: {
                        action: 'previousYear'
                      }
                    }
                  ]
                }
              ],
              verticalContentAlignment: 'Center'
            },
            {
              type: 'Column',
              width: 'stretch',
              items: [
                {
                  type: 'TextBlock',
                  text: '${selectedYear}',
                  size: 'Medium',
                  weight: 'Bolder',
                  horizontalAlignment: 'Center'
                }
              ],
              verticalContentAlignment: 'Center'
            },
            {
              type: 'Column',
              width: 'auto',
              items: [
                {
                  type: 'ActionSet',
                  actions: [
                    {
                      type: 'Action.Submit',
                      title: '${nextYearLabel}',
                      data: {
                        action: 'nextYear'
                      }
                    }
                  ]
                }
              ],
              verticalContentAlignment: 'Center'
            }
          ]
        },
        // Loading indicator
        {
          type: 'TextBlock',
          text: 'Loading holidays...',
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
        // Holiday list
        {
          type: 'Container',
          items: [
            {
              type: 'Container',
              $data: '${holidays}',
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
                          style: 'Default'
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
                        }
                      ],
                      verticalContentAlignment: 'Center'
                    }
                  ]
                }
              ]
            }
          ],
          $when: '${!isLoading && errorMessage == null}'
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

  public onAction(action: IActionArguments): void {
    if (action.type !== 'Submit') {
      return;
    }

    const data = action.data as { action: string };
    const ace = this.context as unknown as UpcomingEventAdaptiveCardExtension;

    if (data.action === 'previousYear') {
      ace.setSelectedYear(this.state.selectedYear - 1);
      this._loadHolidays().then(() => {
        this.quickViewNavigator.replace('UpcomingEvent_HOLIDAY_QUICK_VIEW');
      });
    } else if (data.action === 'nextYear') {
      ace.setSelectedYear(this.state.selectedYear + 1);
      this._loadHolidays().then(() => {
        this.quickViewNavigator.replace('UpcomingEvent_HOLIDAY_QUICK_VIEW');
      });
    }
  }
}
