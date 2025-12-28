# Holidays & Birthdays Web Part

An SPFx React web part for SharePoint Online that displays upcoming holidays and birthdays from a SharePoint list. Features include annual recurrence for birthdays and fixed-date holidays, support for moving holidays (like Labor Day), and expandable views (30-day default, 365-day expanded).

## Features

- **Automatic List Provisioning**: Creates the required SharePoint list and fields automatically on first run (idempotent)
- **Recurrence Support**: 
  - Annual recurrence for birthdays and fixed-date holidays
  - Complex recurrence rules for moving holidays (e.g., "1st Monday of September" for Labor Day)
- **Flexible Display**:
  - Default view: Next 30 days
  - Expandable view: Next 365 days
  - Grouped by month
  - Optional images and type badges
- **User-Friendly UI**: Built with Fluent UI, responsive design, empty states, and error handling

## Prerequisites

- Node.js (v16.13.0 or v18.x)
- npm or yarn
- SharePoint Online tenant
- SPFx development environment configured

## Building the Solution

1. **Install dependencies**:
   ```bash
   npm install
   ```

2. **Build for production**:
   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

   This creates a `.sppkg` file in the `solution` folder.

## Deployment

### Deploy to Site Collection App Catalog

1. Navigate to your SharePoint site
2. Go to Site Contents → Site App Catalog (or your Site Collection App Catalog)
3. Upload the `.sppkg` file from `solution/holidays-birthdays-webpart.sppkg`
4. When prompted, check "Make this solution available to all sites in the organization" (or leave unchecked for site-scoped deployment)
5. Click **Deploy**

### Add Web Part to a Page

1. Edit a SharePoint page (Modern page)
2. Click the **+** button to add a web part
3. Find **"Holidays & Birthdays"** in the web part picker
4. Add it to your page
5. Configure web part properties if needed:
   - **Default Days**: Number of days for default view (default: 30)
   - **Expanded Days**: Number of days for expanded view (default: 365)
   - **List Title**: Name of the SharePoint list (default: "HolidaysAndBirthdays")
   - **Show Images**: Toggle to show/hide event images
   - **Show Type Badges**: Toggle to show/hide event type badges
   - **Allow List Provisioning**: Enable automatic list creation (default: true)

## List Schema

The web part uses a SharePoint list with the following fields:

- **Title** (default field): Display name of the event
- **EventDate** (Date and Time): The base date for the event
- **EventType** (Choice): "Birthday" or "Holiday"
- **IsAnnualRecurrence** (Yes/No): Whether the event recurs annually (default: Yes)
- **RecurrenceRule** (Single line of text): Optional recurrence pattern
  - Format: `MONTHLY_BY_NTH_WEEKDAY:MM:WEEKDAY:NTH`
  - Example: `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1` for Labor Day (1st Monday of September)
- **ImageUrl** (Hyperlink/Picture): Optional image URL for the event
- **Notes** (Multiple lines of text): Optional notes/description

## Adding Events

### Simple Annual Recurrence (Birthdays, Fixed Holidays)

1. Create a new item in the list
2. Set **Title**: e.g., "New Year's Day"
3. Set **EventDate**: e.g., January 1, 2024 (the year doesn't matter for annual recurrence)
4. Set **EventType**: "Holiday" or "Birthday"
5. Set **IsAnnualRecurrence**: Yes (default)
6. Leave **RecurrenceRule** empty

### Moving Holidays (e.g., Labor Day)

1. Create a new item
2. Set **Title**: "Labor Day"
3. Set **EventDate**: Any date in September (e.g., September 1, 2024)
4. Set **EventType**: "Holiday"
5. Set **IsAnnualRecurrence**: Yes
6. Set **RecurrenceRule**: `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1`
   - `09` = September (month number)
   - `MONDAY` = Day of week
   - `1` = First occurrence (1st, 2nd, 3rd, 4th, or last)

## Permissions Required

- **To create the list automatically**: User must have "Manage Lists" permission (typically Site Members or higher)
- **To view events**: User must have "Read" permission on the list (typically all site users)
- **To add/edit events**: User must have "Contribute" permission on the list

## Troubleshooting

### List Not Found Error

- Ensure the list name matches exactly (case-sensitive)
- Check if "Allow List Provisioning" is enabled in web part settings
- Verify user has permissions to create lists

### Events Not Showing

- Verify items exist in the list
- Check that EventDate is set correctly
- Ensure IsAnnualRecurrence is set to Yes for recurring events
- Verify RecurrenceRule format is correct (if using)

### Permission Errors

- Contact your SharePoint administrator to grant appropriate permissions
- Ensure the web part is deployed to an app catalog the site can access

## Development

To work on this project locally:

1. Clone the repository
2. Run `npm install`
3. Update `config/serve.json` with your tenant URL
4. Run `gulp serve` to test in the local workbench
5. Or deploy to your tenant and test on a real page

## License

This project is provided as-is for use in SharePoint Online environments.

