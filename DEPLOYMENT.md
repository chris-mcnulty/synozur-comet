# Holidays & Birthdays Web Part - Deployment Guide

## Quick Start

### 1. Build the Solution

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` file will be created in the `solution` folder.

### 2. Deploy to Site Collection App Catalog

1. Navigate to your SharePoint site
2. Go to **Site Contents** → **Site App Catalog** (or Site Collection App Catalog)
3. Click **Upload** or drag-and-drop the `.sppkg` file from `solution/holidays-birthdays-webpart.sppkg`
4. Click **Deploy** when prompted
5. Optionally check "Make this solution available to all sites in the organization" if deploying tenant-wide

### 3. Add to a Page

1. Edit any modern SharePoint page
2. Click the **+** button to add a web part
3. Find **"Holidays & Birthdays"** in the web part picker
4. The web part will automatically create the list on first use (if provisioning is enabled)

## List Schema

The web part uses these fields:

| Field Name | Type | Required | Default | Description |
|------------|------|----------|---------|-------------|
| Title | Single line of text | Yes | - | Event name |
| EventDate | Date and Time | Yes | - | Base date for recurrence |
| EventType | Choice | Yes | Birthday | "Birthday" or "Holiday" |
| IsAnnualRecurrence | Yes/No | No | Yes | Whether event recurs annually |
| RecurrenceRule | Single line of text | No | - | Complex recurrence pattern (e.g., Labor Day) |
| ImageUrl | Hyperlink/Picture | No | - | Optional image URL |
| Notes | Multiple lines of text | No | - | Optional notes |

## Example List Items

### Birthday (Annual Recurrence)

- **Title**: John Doe
- **EventDate**: 1985-03-15 (any year works)
- **EventType**: Birthday
- **IsAnnualRecurrence**: Yes
- **RecurrenceRule**: (leave empty)

### Fixed Holiday (New Year's Day)

- **Title**: New Year's Day
- **EventDate**: 2024-01-01 (any year works)
- **EventType**: Holiday
- **IsAnnualRecurrence**: Yes
- **RecurrenceRule**: (leave empty)

### Moving Holiday (Labor Day - 1st Monday of September)

- **Title**: Labor Day
- **EventDate**: 2024-09-01 (any date in September)
- **EventType**: Holiday
- **IsAnnualRecurrence**: Yes
- **RecurrenceRule**: `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1`

**RecurrenceRule Format**: `MONTHLY_BY_NTH_WEEKDAY:MM:WEEKDAY:NTH`
- `MM`: Month number (01-12)
- `WEEKDAY`: Day name (SUNDAY, MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY)
- `NTH`: Occurrence number (1=first, 2=second, 3=third, 4=fourth, 5=last)
- **Note**: Use `5` for the last occurrence of a weekday in the month (e.g., last Monday of May for Memorial Day)

### Other Examples

- **Thanksgiving** (4th Thursday of November): `MONTHLY_BY_NTH_WEEKDAY:11:THURSDAY:4`
- **Memorial Day** (last Monday of May): `MONTHLY_BY_NTH_WEEKDAY:05:MONDAY:5`

## Permissions

- **List Creation**: Requires "Manage Lists" permission (Site Members or higher)
- **Viewing Events**: Requires "Read" permission on the list
- **Adding/Editing Events**: Requires "Contribute" permission on the list

## Troubleshooting

### "List not found" Error

- Verify list name matches exactly (case-sensitive)
- Enable "Allow List Provisioning" in web part settings
- Check user permissions to create lists

### Events Not Appearing

- Verify items exist in the list
- Check EventDate format (YYYY-MM-DD)
- Ensure IsAnnualRecurrence is Yes for recurring events
- Verify RecurrenceRule format if using moving holidays

### Permission Errors

- Contact SharePoint administrator
- Ensure web part is deployed from accessible app catalog
- Verify user has appropriate site permissions

## Development

For local development:

```bash
npm install
gulp serve
```

Update `config/serve.json` with your tenant URL for testing.

