import { spfi, SPFI } from '@pnp/sp';
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
  private _sp: SPFI;
  private _listTitle: string;

  constructor(sp: SPFI, listTitle: string) {
    this._sp = sp;
    this._listTitle = listTitle;
  }

  public async ensureListExists(): Promise<IProvisioningResult> {
    try {
      // Check if list exists
      let listExists = false;
      let list: any;
      
      try {
        list = this._sp.web.lists.getByTitle(this._listTitle);
        await list.select('Id')();
        listExists = true;
      } catch (err) {
        // List doesn't exist, create it
        listExists = false;
      }

      if (!listExists) {
        // Create the list
        const listAddResult = await this._sp.web.lists.add(this._listTitle, "Holidays and Birthdays calendar", 100);
        // Get the list ID from the result (PnPjs returns { Id: ... })
        const listId = (listAddResult as any).Id || (listAddResult as any).data?.Id;
        if (listId) {
          list = this._sp.web.lists.getById(listId);
        } else {
          // Fallback to get by title
          list = this._sp.web.lists.getByTitle(this._listTitle);
        }
        listExists = true;
      } else {
        list = this._sp.web.lists.getByTitle(this._listTitle);
      }

      // Ensure all required fields exist
      await this.ensureFields(list);

      return {
        success: true,
        listExists: listExists
      };
    } catch (error) {
      console.error('List provisioning error:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Unknown error',
        listExists: false
      };
    }
  }

  private async ensureFields(list: any): Promise<void> {
    const fieldsToCreate = [
      {
        internalName: 'EventDate',
        displayName: 'Event Date',
        type: 'DateTime',
        required: true,
        format: 'DateOnly'
      },
      {
        internalName: 'EventType',
        displayName: 'Event Type',
        type: 'Choice',
        choices: ['Birthday', 'Holiday'],
        required: true,
        defaultValue: 'Birthday'
      },
      {
        internalName: 'IsAnnualRecurrence',
        displayName: 'Is Annual Recurrence',
        type: 'Boolean',
        required: false,
        defaultValue: true
      },
      {
        internalName: 'RecurrenceRule',
        displayName: 'Recurrence Rule',
        type: 'Text',
        required: false
      },
      {
        internalName: 'ImageUrl',
        displayName: 'Image URL',
        type: 'URL',
        required: false
      },
      {
        internalName: 'Notes',
        displayName: 'Notes',
        type: 'Note',
        required: false,
        richText: false
      }
    ];

    const existingFields = await list.fields.select('InternalName')();
    const existingFieldNames = existingFields.map((f: any) => f.InternalName);

    for (const fieldDef of fieldsToCreate) {
      if (!existingFieldNames.includes(fieldDef.internalName)) {
        try {
          await this.createField(list, fieldDef);
        } catch (error) {
          console.warn(`Failed to create field ${fieldDef.internalName}:`, error);
        }
      }
    }
  }

  private async createField(list: any, fieldDef: any): Promise<void> {
    const fieldXml = this.getFieldXml(fieldDef);
    
    await list.fields.createFieldAsXml(fieldXml);
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

