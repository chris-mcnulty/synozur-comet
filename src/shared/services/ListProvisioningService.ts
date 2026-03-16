import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";

export interface IProvisioningResult {
  success: boolean;
  error?: string;
  listExists: boolean;
}

export class ListProvisioningService {
  private _listTitle: string;

  // Default seed events for a newly provisioned list
  private readonly _defaultEvents = [
    {
      Title: "New Year's Day",
      EventDate: new Date('2026-01-01'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: '',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Martin Luther King Jr. Day',
      EventDate: new Date('2026-01-19'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: '',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Presidents Day',
      EventDate: new Date('2026-02-21'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: 'MONTHLY_BY_NTH_WEEKDAY:02:MONDAY:3',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Memorial Day',
      EventDate: new Date('2026-05-26'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: 'MONTHLY_BY_NTH_WEEKDAY:05:MONDAY:5',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Independence Day',
      EventDate: new Date('2026-07-04'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: '',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Labor Day',
      EventDate: new Date('2026-09-07'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: 'MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1',
      ImageUrl: '',
      Active: true
    },
    {
      Title: "Indigenous People's Day",
      EventDate: new Date('2026-10-12'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: 'MONTHLY_BY_NTH_WEEKDAY:10:MONDAY:2',
      ImageUrl: '',
      Active: true
    },
    {
      Title: "Veterans' Day",
      EventDate: new Date('2026-11-11'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: '',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Thanksgiving',
      EventDate: new Date('2026-11-26'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: 'MONTHLY_BY_NTH_WEEKDAY:11:THURSDAY:4',
      ImageUrl: '',
      Active: true
    },
    {
      Title: 'Christmas',
      EventDate: new Date('2026-12-25'),
      EventType: 'Holiday',
      IsAnnualRecurrence: true,
      RecurrenceRule: '',
      ImageUrl: '',
      Active: true
    }
  ];

  constructor(listTitle: string) {
    this._listTitle = listTitle;
  }

  public async ensureListExists(): Promise<IProvisioningResult> {
    try {
      // Check if list exists
      let listExists = false;
      let list: any;
      
      try {
        list = sp.web.lists.getByTitle(this._listTitle);
        await list.select('Id')();
        listExists = true;
      } catch (err: any) {
        // List doesn't exist if we get a 404 or ItemNotFoundException
        const errorMsg = err?.message || '';
        if (errorMsg.includes('404') || errorMsg.includes('does not exist') || errorMsg.includes('ItemNotFoundException')) {
          listExists = false;
        } else {
          // Some other error checking for the list - propagate it
          throw err;
        }
      }

      if (!listExists) {
        // Create the list (template 100 = Generic List)
        try {
          const listAddResult = await sp.web.lists.add(this._listTitle, "Holidays and Birthdays calendar", 100);
          // Get the list ID from the result (PnPjs v2 returns { Id: ... })
          const listId = (listAddResult as any).Id || (listAddResult as any).data?.Id;
          if (listId) {
            list = sp.web.lists.getById(listId);
          } else {
            // Fallback to get by title - wait a moment for list to be available
            await new Promise(resolve => setTimeout(resolve, 500));
            list = sp.web.lists.getByTitle(this._listTitle);
          }
          listExists = true;
          // Wait a bit more before creating fields to ensure list is fully initialized
          await new Promise(resolve => setTimeout(resolve, 1000));
        } catch (createError: any) {
          console.error('Error creating list:', createError);
          const errorMessage = createError?.message || 'Unknown error creating list';
          return {
            success: false,
            error: `Failed to create list: ${errorMessage}. Please ensure you have "Manage Lists" permission.`,
            listExists: false
          };
        }
      } else {
        list = sp.web.lists.getByTitle(this._listTitle);
      }

      // Ensure all required fields exist
      try {
        await this.ensureFields(list);
        // Seed default items if the list is empty
        await this.seedDefaultItems(list);
      } catch (fieldError: any) {
        console.error('Error ensuring fields:', fieldError);
        const fieldErrorMessage = fieldError?.message || 'Unknown error';
        // If fields already exist, that's okay - don't fail
        if (!fieldErrorMessage.includes('already exists') && !fieldErrorMessage.includes('duplicate')) {
          return {
            success: false,
            error: `List exists but failed to create required fields: ${fieldErrorMessage}. Please ensure you have "Manage Lists" permission.`,
            listExists: listExists
          };
        }
      }

      return {
        success: true,
        listExists: listExists
      };
    } catch (error) {
      console.error('List provisioning error:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      return {
        success: false,
        error: errorMessage,
        listExists: false
      };
    }
  }

  private async seedDefaultItems(list: any): Promise<void> {
    try {
      // Check if the list already has items
      const existingItems = await list.items.top(1).select('Id')();
      const hasItems = Array.isArray(existingItems) ? existingItems.length > 0 : (existingItems?.value?.length ?? 0) > 0;

      if (hasItems) {
        return;
      }

      for (const ev of this._defaultEvents) {
        try {
          await list.items.add({
            Title: ev.Title,
            EventDate: ev.EventDate,
            EventType: ev.EventType,
            IsAnnualRecurrence: ev.IsAnnualRecurrence,
            RecurrenceRule: ev.RecurrenceRule,
            ImageUrl: ev.ImageUrl,
            Active: ev.Active
          });
          // Small delay between item creations to avoid throttling
          await new Promise(resolve => setTimeout(resolve, 100));
        } catch (itemError: any) {
          console.error('Error seeding default item:', ev.Title, itemError);
        }
      }
    } catch (error) {
      console.error('Error checking or seeding default items:', error);
    }
  }

  private async ensureFields(list: any): Promise<void> {
    const fieldsToCreate = [
      {
        internalName: 'EventDate',
        displayName: 'EventDate',
        type: 'DateTime',
        required: true,
        format: 'DateOnly'
      },
      {
        internalName: 'EventType',
        displayName: 'EventType',
        type: 'Choice',
        choices: ['Birthday', 'Holiday'],
        required: true,
        defaultValue: 'Birthday'
      },
      {
        internalName: 'IsAnnualRecurrence',
        displayName: 'IsAnnualRecurrence',
        type: 'Boolean',
        required: false,
        defaultValue: true
      },
      {
        internalName: 'RecurrenceRule',
        displayName: 'RecurrenceRule',
        type: 'Text',
        required: false
      },
      {
        internalName: 'ImageUrl',
        displayName: 'ImageUrl',
        type: 'URL',
        required: false
      },
      {
        internalName: 'Notes',
        displayName: 'Notes',
        type: 'Note',
        required: false,
        richText: false
      },
      {
        internalName: 'Active',
        displayName: 'Active',
        type: 'Boolean',
        required: false,
        defaultValue: true
      }
    ];

    try {
      // Get existing fields - handle both array and object responses
      const existingFieldsResponse = await list.fields.select('InternalName')();
      const existingFields = Array.isArray(existingFieldsResponse) ? existingFieldsResponse : existingFieldsResponse.value || [];
      const existingFieldNames = existingFields.map((f: any) => f.InternalName || f.internalName);
      console.log('Existing fields:', existingFieldNames);

      const fieldCreationErrors: string[] = [];

      for (const fieldDef of fieldsToCreate) {
        if (!existingFieldNames.includes(fieldDef.internalName)) {
          try {
            await this.createField(list, fieldDef);
            // Small delay between field creations to avoid throttling
            await new Promise(resolve => setTimeout(resolve, 200));
          } catch (error: any) {
            const errorMsg = error?.message || 'Unknown error';
            // If field already exists, that's okay
            if (!errorMsg.includes('already exists') && !errorMsg.includes('duplicate')) {
              console.error(`Failed to create field ${fieldDef.internalName}:`, error);
              fieldCreationErrors.push(`${fieldDef.displayName} (${errorMsg})`);
            }
          }
        }
      }

      // If there were errors creating fields, throw an error
      if (fieldCreationErrors.length > 0) {
        throw new Error(`Failed to create fields: ${fieldCreationErrors.join('; ')}`);
      }
    } catch (error) {
      // Re-throw to be caught by caller
      throw error;
    }
  }

  private async createField(list: any, fieldDef: any): Promise<void> {
    const fieldXml = this.getFieldXml(fieldDef);
    console.log(`Creating field ${fieldDef.internalName} with XML:`, fieldXml);
    
    try {
      await list.fields.createFieldAsXml(fieldXml);
      console.log(`Successfully created field ${fieldDef.internalName}`);
    } catch (error: any) {
      console.error(`Error creating field ${fieldDef.internalName}:`, error, 'XML:', fieldXml);
      throw error;
    }
  }

  private getFieldXml(fieldDef: any): string {
    switch (fieldDef.type) {
      case 'DateTime':
        return `<Field Type="DateTime" DisplayName="${fieldDef.displayName}" Name="${fieldDef.internalName}" Format="${fieldDef.format || 'DateTime'}" Required="${fieldDef.required || 'FALSE'}" />`;
      
      case 'Choice':
        const choiceOptions = fieldDef.choices.map((c: string) => `<CHOICE>${c}</CHOICE>`).join('');
        return `<Field Type="Choice" DisplayName="${fieldDef.displayName}" Name="${fieldDef.internalName}" Required="${fieldDef.required || 'FALSE'}"><CHOICES>${choiceOptions}</CHOICES><Default>${fieldDef.defaultValue || ''}</Default></Field>`;
      
      case 'Boolean':
        return `<Field Type="Boolean" DisplayName="${fieldDef.displayName}" Name="${fieldDef.internalName}" Required="${fieldDef.required || 'FALSE'}"><Default>${fieldDef.defaultValue === true ? '1' : '0'}</Default></Field>`;
      
      case 'Text':
        return `<Field Type="Text" DisplayName="${fieldDef.displayName}" Name="${fieldDef.internalName}" Required="${fieldDef.required || 'FALSE'}" />`;
      
      case 'URL':
        return `<Field Type="URL" DisplayName="${fieldDef.displayName}" Name="${fieldDef.internalName}" Required="${fieldDef.required || 'FALSE'}" />`;
      
      case 'Note':
        return `<Field Type="Note" DisplayName="${fieldDef.displayName}" Name="${fieldDef.internalName}" Required="${fieldDef.required || 'FALSE'}" RichText="${fieldDef.richText ? 'TRUE' : 'FALSE'}" />`;
      
      default:
        throw new Error(`Unsupported field type: ${fieldDef.type}`);
    }
  }
}
