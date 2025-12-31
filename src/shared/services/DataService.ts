import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IEventOccurrence, RecurrenceCalculator } from './RecurrenceCalculator';

export interface IListItem {
  Id: number;
  Title: string;
  EventDate: string;
  EventType: string;
  IsAnnualRecurrence: boolean;
  RecurrenceRule?: string;
  ImageUrl?: { Url: string; Description: string };
  Notes?: string;
  Active?: boolean;
}

export class DataService {
  private _listTitle: string;

  constructor(listTitle: string) {
    this._listTitle = listTitle;
  }

  /**
   * Get all events within the specified number of days
   */
  public async getEvents(maxDays: number = 365): Promise<IEventOccurrence[]> {
    try {
      const items = await this.fetchListItems();
      const allOccurrences: IEventOccurrence[] = [];

      for (const item of items) {
        const eventDate = new Date(item.EventDate);
        const isAnnualRecurrence = item.IsAnnualRecurrence !== false; // Default to true
        
        // Extract image URL if present
        const imageUrl = this.extractImageUrl(item.ImageUrl);

        const occurrences = RecurrenceCalculator.calculateNextOccurrences(
          eventDate,
          isAnnualRecurrence,
          item.RecurrenceRule || null,
          item.Title,
          item.EventType,
          imageUrl,
          item.Notes,
          maxDays
        );

        allOccurrences.push(...occurrences);
      }

      // Sort by date ascending
      allOccurrences.sort((a, b) => a.date.getTime() - b.date.getTime());

      return allOccurrences;
    } catch (error) {
      console.error('Error fetching events:', error);
      throw error;
    }
  }

  /**
   * Get the next upcoming event of a specific type (for ACE Card View)
   */
  public async getNextEvent(eventType: 'Holiday' | 'Birthday'): Promise<IEventOccurrence | null> {
    try {
      const events = await this.getEvents(365);
      const filtered = events.filter(e => e.eventType === eventType);
      return filtered.length > 0 ? filtered[0] : null;
    } catch (error) {
      console.error('Error fetching next event:', error);
      throw error;
    }
  }

  /**
   * Get the next N birthdays (for Birthday ACE Quick View)
   */
  public async getNextBirthdays(count: number = 5): Promise<IEventOccurrence[]> {
    try {
      const events = await this.getEvents(365);
      return events
        .filter(e => e.eventType === 'Birthday')
        .slice(0, count);
    } catch (error) {
      console.error('Error fetching next birthdays:', error);
      throw error;
    }
  }

  /**
   * Get all holidays for a specific year (for Holiday ACE Quick View)
   */
  public async getHolidaysForYear(year: number): Promise<IEventOccurrence[]> {
    try {
      const items = await this.fetchListItems();
      const allOccurrences: IEventOccurrence[] = [];

      for (const item of items) {
        // Only process holidays
        if (item.EventType !== 'Holiday') {
          continue;
        }

        const eventDate = new Date(item.EventDate);
        const isAnnualRecurrence = item.IsAnnualRecurrence !== false;
        const imageUrl = this.extractImageUrl(item.ImageUrl);

        const occurrences = RecurrenceCalculator.calculateOccurrencesForYear(
          eventDate,
          isAnnualRecurrence,
          item.RecurrenceRule || null,
          item.Title,
          item.EventType,
          imageUrl,
          item.Notes,
          year
        );

        allOccurrences.push(...occurrences);
      }

      // Sort by date ascending
      allOccurrences.sort((a, b) => a.date.getTime() - b.date.getTime());

      return allOccurrences;
    } catch (error) {
      console.error('Error fetching holidays for year:', error);
      throw error;
    }
  }

  /**
   * Fetch all active list items from SharePoint
   */
  private async fetchListItems(): Promise<IListItem[]> {
    const list = sp.web.lists.getByTitle(this._listTitle);
    
    const items: IListItem[] = await list.items.select(
      'Id',
      'Title',
      'EventDate',
      'EventType',
      'IsAnnualRecurrence',
      'RecurrenceRule',
      'ImageUrl',
      'Notes',
      'Active'
    )();

    // Filter to only include active entries (default to true if Active field is not set)
    return items.filter(item => item.Active !== false);
  }

  /**
   * Extract image URL from various formats
   */
  private extractImageUrl(imageUrl: { Url: string; Description: string } | string | undefined): string | undefined {
    if (!imageUrl) {
      return undefined;
    }
    return typeof imageUrl === 'string' ? imageUrl : imageUrl.Url;
  }
}
