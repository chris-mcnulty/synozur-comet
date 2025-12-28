import { spfi, SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IEventOccurrence } from './RecurrenceCalculator';
import { RecurrenceCalculator } from './RecurrenceCalculator';

export interface IListItem {
  Id: number;
  Title: string;
  EventDate: string;
  EventType: string;
  IsAnnualRecurrence: boolean;
  RecurrenceRule?: string;
  ImageUrl?: { Url: string; Description: string };
  Notes?: string;
}

export class DataService {
  private _sp: SPFI;
  private _listTitle: string;

  constructor(sp: SPFI, listTitle: string) {
    this._sp = sp;
    this._listTitle = listTitle;
  }

  public async getEvents(maxDays: number = 365): Promise<IEventOccurrence[]> {
    try {
      const list = this._sp.web.lists.getByTitle(this._listTitle);
      
      // Get all items from the list
      const items: IListItem[] = await list.items.select(
        'Id',
        'Title',
        'EventDate',
        'EventType',
        'IsAnnualRecurrence',
        'RecurrenceRule',
        'ImageUrl',
        'Notes'
      )();

      const allOccurrences: IEventOccurrence[] = [];

      for (const item of items) {
        const eventDate = new Date(item.EventDate);
        const isAnnualRecurrence = item.IsAnnualRecurrence !== false; // Default to true
        
        // Extract image URL if present
        let imageUrl: string | undefined;
        if (item.ImageUrl) {
          imageUrl = typeof item.ImageUrl === 'string' ? item.ImageUrl : item.ImageUrl.Url;
        }

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
}

