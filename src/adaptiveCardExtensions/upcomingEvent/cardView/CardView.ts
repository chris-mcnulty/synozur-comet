import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'UpcomingEventAdaptiveCardExtensionStrings';
import {
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState
} from '../IUpcomingEventAdaptiveCardExtensionProps';
import {
  HOLIDAY_QUICK_VIEW_REGISTRY_ID,
  BIRTHDAY_QUICK_VIEW_REGISTRY_ID
} from '../UpcomingEventAdaptiveCardExtension';

export class CardView extends BaseImageCardView<
  IUpcomingEventAdaptiveCardExtensionProps,
  IUpcomingEventAdaptiveCardExtensionState
> {
  /**
   * Buttons will not be visible if card size is 'Small'
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    const quickViewId = this.properties.displayMode === 'Birthday'
      ? BIRTHDAY_QUICK_VIEW_REGISTRY_ID
      : HOLIDAY_QUICK_VIEW_REGISTRY_ID;

    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: quickViewId
          }
        }
      }
    ];
  }

  public get data(): IImageCardParameters {
    const { nextEventTitle, nextEventDate, nextEventImageUrl, daysUntil, isLoading } = this.state;
    const { displayMode, showDaysUntil, showBranding } = this.properties;

    // Type label
    const typeLabel = displayMode === 'Birthday' ? 'BIRTHDAY' : 'HOLIDAY';

    // Format the date
    let dateText = '';
    if (nextEventDate) {
      dateText = nextEventDate.toLocaleDateString('en-US', {
        weekday: 'short',
        month: 'short',
        day: 'numeric'
      });
    }

    // Days until text
    let daysUntilText = '';
    if (showDaysUntil && nextEventDate) {
      if (daysUntil === 0) {
        daysUntilText = strings.DaysUntilToday;
      } else if (daysUntil === 1) {
        daysUntilText = strings.DaysUntilTomorrow;
      } else {
        daysUntilText = strings.DaysUntilFormat.replace('{0}', daysUntil.toString());
      }
    }

    // Build primary text with type label
    let primaryText = isLoading ? strings.LoadingMessage : `${typeLabel}: ${nextEventTitle}`;

    // Build secondary text
    let secondaryText = dateText;
    if (daysUntilText) {
      secondaryText = secondaryText ? `${secondaryText} • ${daysUntilText}` : daysUntilText;
    }

    // Add branding if enabled
    if (showBranding) {
      secondaryText = secondaryText ? `${secondaryText}\n${strings.CompanyName}` : strings.CompanyName;
    }

    return {
      primaryText: primaryText,
      imageUrl: nextEventImageUrl,
      title: displayMode === 'Birthday' ? strings.BirthdaysTitle : strings.HolidaysTitle,
      secondaryText: secondaryText
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    const quickViewId = this.properties.displayMode === 'Birthday'
      ? BIRTHDAY_QUICK_VIEW_REGISTRY_ID
      : HOLIDAY_QUICK_VIEW_REGISTRY_ID;

    return {
      type: 'QuickView',
      parameters: {
        view: quickViewId
      }
    };
  }
}
