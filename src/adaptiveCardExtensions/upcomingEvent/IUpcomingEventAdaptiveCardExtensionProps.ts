export type ACEDisplayMode = 'Holiday' | 'Birthday';

export interface IUpcomingEventAdaptiveCardExtensionProps {
  /**
   * Display mode: Show holidays or birthdays
   */
  displayMode: ACEDisplayMode;

  /**
   * SharePoint list title containing events
   */
  listTitle: string;

  /**
   * Show "in X days" countdown on cards
   */
  showDaysUntil: boolean;

  /**
   * Show Synozur branding footer on cards
   */
  showBranding: boolean;

  /**
   * Number of birthdays to show in Birthday Quick View (3-10)
   */
  birthdayCount: number;
}

export interface IUpcomingEventAdaptiveCardExtensionState {
  /**
   * Title of the next event
   */
  nextEventTitle: string;

  /**
   * Date of the next event
   */
  nextEventDate: Date | null;

  /**
   * Image URL for the next event
   */
  nextEventImageUrl: string;

  /**
   * Days until the next event
   */
  daysUntil: number;

  /**
   * Whether data is loading
   */
  isLoading: boolean;

  /**
   * Error message if any
   */
  error: string | null;

  /**
   * Current year for holiday quick view navigation
   */
  selectedYear: number;
}
