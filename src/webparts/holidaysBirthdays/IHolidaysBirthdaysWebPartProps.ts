export type DisplayMode = 'Both' | 'Holidays' | 'Birthdays';

export interface IHolidaysBirthdaysWebPartProps {
  daysDefault: number;
  daysExpanded: number;
  listTitle: string;
  showImages: boolean;
  showTypeBadges: boolean;
  allowListProvisioning: boolean;
  displayMode: DisplayMode;
}

