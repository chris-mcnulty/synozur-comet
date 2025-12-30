# Holidays & Birthdays Web Part - User Guide

This guide is for SharePoint administrators and end users who need to deploy, configure, and update the Holidays & Birthdays web part. No development knowledge required.

## Table of Contents

1. [Deploying the Web Part](#deploying-the-web-part)
2. [Adding the Web Part to a Page](#adding-the-web-part-to-a-page)
3. [Configuring the Web Part](#configuring-the-web-part)
4. [Setting Up Events](#setting-up-events)
5. [Common Holiday Recurrence Patterns](#common-holiday-recurrence-patterns)
6. [Updating the Web Part](#updating-the-web-part)
7. [Troubleshooting](#troubleshooting)

---

## Deploying the Web Part

### Prerequisites

- SharePoint Administrator permissions (for tenant-wide deployment)
- Site Collection Administrator permissions (for site collection deployment)
- The `.sppkg` file provided by your development team

### Option 1: Tenant-Wide Deployment (Recommended)

This makes the web part available to all sites in your organization.

1. **Navigate to SharePoint Admin Center**
   - Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
   - Click **Admin centers** → **SharePoint**
   - Or go directly to: `https://[your-tenant]-admin.sharepoint.com`

2. **Access App Catalog**
   - In the left navigation, click **More features**
   - Under **Apps**, click **App Catalog**
   - Click **Apps for SharePoint**

3. **Upload the Package**
   - Click **Upload** or drag-and-drop the `.sppkg` file
   - When prompted, check **"Replace the existing file if it exists"** (if updating)
   - Click **OK**

4. **Deploy the Solution**
   - Click **Deploy** when prompted
   - **Important**: Check **"Make this solution available to all sites in the organization"**
   - Click **OK**
   - Wait for the deployment to complete (usually 1-2 minutes)

5. **Verify Deployment**
   - The solution should show as "Deployed" with a green checkmark
   - The web part is now available to all sites in your tenant

### Option 2: Site Collection Deployment

This makes the web part available only to a specific site collection.

1. **Navigate to Your Site**
   - Go to the SharePoint site where you want to deploy the web part

2. **Access Site App Catalog**
   - Click the **Settings** gear icon (top right)
   - Click **Site contents**
   - Look for **"Apps for SharePoint"** or **"Site App Catalog"**
   - If you don't see it, you may need to enable it in Site Settings

3. **Upload and Deploy**
   - Click **Upload** or drag-and-drop the `.sppkg` file
   - Check **"Replace the existing file if it exists"** (if updating)
   - Click **Deploy**
   - The web part is now available to this site collection

---

## Adding the Web Part to a Page

1. **Edit a SharePoint Page**
   - Navigate to any modern SharePoint page
   - Click **Edit** (top right)

2. **Add the Web Part**
   - Click the **+** button where you want to add the web part
   - In the web part picker, search for **"Holidays & Birthdays"**
   - Click on **"Holidays & Birthdays"** to add it to the page

3. **Save the Page**
   - Click **Publish** or **Republish** to save your changes

---

## Configuring the Web Part

After adding the web part to a page, you can configure its settings.

1. **Open Property Pane**
   - Click on the web part to select it
   - Click the **Edit web part** icon (pencil icon) that appears

2. **Configure Settings**

   - **Default Days (30-day view)**: Number of days to show in the default view (default: 30)
   - **Expanded Days (365-day view)**: Number of days to show when expanded (default: 365)
   - **List Title**: Name of the SharePoint list that stores events (default: "HolidaysAndBirthdays")
   - **Show Images**: Toggle to show/hide event images (default: On)
   - **Show Type Badges**: Toggle to show/hide "Holiday" or "Birthday" badges (default: On)
   - **Display Mode**: Choose what to display
     - **Holidays & Birthdays**: Show both (default)
     - **Holidays Only**: Show only holidays
     - **Birthdays Only**: Show only birthdays
   - **Show Footer**: Toggle to show/hide the footer with copyright information (default: On)
   - **Allow List Provisioning**: Enable automatic list creation (default: On)
     - **Important**: Users need "Manage Lists" permission for this to work

3. **Save Configuration**
   - Click **OK** or click outside the property pane
   - The web part will automatically refresh with your settings

---

## Setting Up Events

The web part displays events from a SharePoint list. You can either let the web part create the list automatically (if provisioning is enabled) or create it manually.

### Automatic List Creation (Recommended)

If **"Allow List Provisioning"** is enabled:

1. The list is created automatically when the web part first loads
2. Users with "Manage Lists" permission can add events immediately
3. No manual setup required

### Manual List Creation

If you prefer to create the list manually:

1. **Create a New List**
   - Go to **Site Contents** → **New** → **List**
   - Name it: **HolidaysAndBirthdays** (or match your web part's list title setting)
   - Click **Create**

2. **Add Required Fields**
   The list needs these fields (the web part can add them automatically, or you can add manually):

   | Field Name | Type | Required | Default Value |
   |------------|------|----------|---------------|
   | Title | Single line of text | Yes | - |
   | EventDate | Date and Time | Yes | - |
   | EventType | Choice | Yes | Birthday |
   | IsAnnualRecurrence | Yes/No | No | Yes |
   | RecurrenceRule | Single line of text | No | - |
   | ImageUrl | Hyperlink/Picture | No | - |
   | Notes | Multiple lines of text | No | - |

3. **Configure EventType Choice Field**
   - Edit the EventType field
   - Add choices: **Birthday** and **Holiday**
   - Set default to **Birthday**

### Adding Events

1. **Open the List**
   - Go to **Site Contents** → Click on your list (e.g., "HolidaysAndBirthdays")

2. **Add New Item**
   - Click **New** or **+ New**
   - Fill in the fields:
     - **Title**: Name of the event (e.g., "New Year's Day" or "John Doe")
     - **EventDate**: The date of the event (year doesn't matter for annual events)
     - **EventType**: Choose "Holiday" or "Birthday"
     - **IsAnnualRecurrence**: Set to "Yes" for recurring events
     - **RecurrenceRule**: Leave empty for fixed dates, or use patterns for moving holidays (see below)
     - **ImageUrl**: Optional - URL to an image for this event
     - **Notes**: Optional - Additional information about the event
   - Click **Save**

3. **View in Web Part**
   - Return to the page with the web part
   - The event should appear automatically (refresh the page if needed)

---

## Common Holiday Recurrence Patterns

For holidays that fall on the same date every year (like Independence Day on July 4th), leave the **RecurrenceRule** field empty.

For holidays that move each year (like Thanksgiving), use the patterns below.

### Recurrence Rule Format

`MONTHLY_BY_NTH_WEEKDAY:MM:WEEKDAY:NTH`

Where:
- **MM**: Month number (01-12)
- **WEEKDAY**: Day name (MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY)
- **NTH**: Occurrence number
  - `1` = First occurrence
  - `2` = Second occurrence
  - `3` = Third occurrence
  - `4` = Fourth occurrence
  - `5` = Last occurrence

### Common U.S. Federal Holidays

| Holiday | Date Pattern | RecurrenceRule | EventDate Example |
|---------|-------------|----------------|-------------------|
| **New Year's Day** | January 1 (Fixed) | *(leave empty)* | 2024-01-01 |
| **Martin Luther King Jr. Day** | 3rd Monday in January | `MONTHLY_BY_NTH_WEEKDAY:01:MONDAY:3` | 2024-01-01 |
| **Presidents' Day** | 3rd Monday in February | `MONTHLY_BY_NTH_WEEKDAY:02:MONDAY:3` | 2024-02-01 |
| **Memorial Day** | Last Monday in May | `MONTHLY_BY_NTH_WEEKDAY:05:MONDAY:5` | 2024-05-01 |
| **Independence Day** | July 4 (Fixed) | *(leave empty)* | 2024-07-04 |
| **Labor Day** | 1st Monday in September | `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1` | 2024-09-01 |
| **Columbus Day** | 2nd Monday in October | `MONTHLY_BY_NTH_WEEKDAY:10:MONDAY:2` | 2024-10-01 |
| **Veterans Day** | November 11 (Fixed) | *(leave empty)* | 2024-11-11 |
| **Thanksgiving Day** | 4th Thursday in November | `MONTHLY_BY_NTH_WEEKDAY:11:THURSDAY:4` | 2024-11-01 |
| **Christmas Day** | December 25 (Fixed) | *(leave empty)* | 2024-12-25 |

### Step-by-Step: Adding a Moving Holiday

**Example: Adding Labor Day**

1. Open your events list
2. Click **New** to add a new item
3. Fill in:
   - **Title**: `Labor Day`
   - **EventDate**: `2024-09-01` (any date in September works)
   - **EventType**: `Holiday`
   - **IsAnnualRecurrence**: `Yes`
   - **RecurrenceRule**: `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1`
   - **ImageUrl**: *(optional)*
   - **Notes**: *(optional)*
4. Click **Save**

The web part will automatically calculate Labor Day for each year (always the 1st Monday in September).

### Step-by-Step: Adding a Fixed Date Holiday

**Example: Adding Independence Day**

1. Open your events list
2. Click **New** to add a new item
3. Fill in:
   - **Title**: `Independence Day`
   - **EventDate**: `2024-07-04` (the actual date)
   - **EventType**: `Holiday`
   - **IsAnnualRecurrence**: `Yes`
   - **RecurrenceRule**: *(leave empty)*
   - **ImageUrl**: *(optional)*
   - **Notes**: *(optional)*
4. Click **Save**

The web part will show Independence Day on July 4th every year.

### Step-by-Step: Adding a Birthday

1. Open your events list
2. Click **New** to add a new item
3. Fill in:
   - **Title**: `John Doe` (person's name)
   - **EventDate**: `1985-03-15` (their birth date - year doesn't matter)
   - **EventType**: `Birthday`
   - **IsAnnualRecurrence**: `Yes`
   - **RecurrenceRule**: *(leave empty)*
   - **ImageUrl**: *(optional - URL to their photo)*
   - **Notes**: *(optional)*
4. Click **Save**

The web part will show their birthday on March 15th every year.

---

## Updating the Web Part

When a new version of the web part is available, follow these steps to update it.

### Pre-Update Checklist

- [ ] Notify users that the web part will be updated
- [ ] Verify you have the new `.sppkg` file
- [ ] Check the version number of the new package (if provided)

### Update Steps

1. **Navigate to App Catalog**
   - Go to **SharePoint Admin Center** → **More features** → **Apps** → **App Catalog**
   - Click **Apps for SharePoint**

2. **Upload New Version**
   - Click **Upload** or drag-and-drop the new `.sppkg` file
   - **Important**: Check **"Replace the existing file if it exists"**
   - Click **OK**

3. **Deploy the Update**
   - Click **Deploy** when prompted
   - Check **"Make this solution available to all sites in the organization"** (if tenant-wide)
   - Click **OK**
   - Wait for deployment to complete (usually 1-2 minutes)

4. **Verify the Update**
   - The solution should show as "Deployed"
   - Check the version number matches the new version

### Post-Update Steps

1. **Clear Browser Cache** (Recommended)
   - Users should clear their browser cache or do a hard refresh (Ctrl+F5)
   - This ensures they see the latest version

2. **Test the Web Part**
   - Visit a page with the web part
   - Verify it loads correctly
   - Check for any new features or changes

3. **Wait for Propagation**
   - Changes may take 5-15 minutes to appear across all sites
   - CDN caching can take up to 24 hours for full global propagation

### Important Notes About Updates

- **No Downtime**: Updates deploy without interrupting existing web parts
- **Automatic**: All web part instances automatically use the new version
- **Property Settings**: Existing web parts keep their current settings
- **Breaking Changes**: If properties changed, you may need to reconfigure web parts

---

## Troubleshooting

### Web Part Not Appearing in Web Part Picker

**Possible Causes:**
- Web part not deployed to App Catalog
- Solution not marked as "Available"
- Browser cache needs clearing

**Solutions:**
1. Verify deployment completed successfully in App Catalog
2. Check that solution shows as "Deployed" with green checkmark
3. Clear browser cache and try again
4. Wait 5-10 minutes for propagation

### "List not found" Error

**Possible Causes:**
- List doesn't exist
- List name doesn't match web part configuration
- User doesn't have permissions

**Solutions:**
1. Enable **"Allow List Provisioning"** in web part settings (if you have permissions)
2. Verify the **List Title** in web part settings matches your list name exactly (case-sensitive)
3. Create the list manually if provisioning is disabled
4. Check user has "Read" permission on the list

### Events Not Appearing

**Possible Causes:**
- No items in the list
- Date format incorrect
- Recurrence settings incorrect
- Display mode filtering events

**Solutions:**
1. Verify items exist in the list
2. Check **EventDate** is in correct format (YYYY-MM-DD)
3. Ensure **IsAnnualRecurrence** is set to "Yes" for recurring events
4. Verify **RecurrenceRule** format is correct (if using moving holidays)
5. Check **Display Mode** setting - make sure it's set to show the event type you added

### Web Part Shows Old Version After Update

**Possible Causes:**
- Browser cache
- CDN caching
- Multiple versions in App Catalog

**Solutions:**
1. Clear browser cache (Ctrl+F5 or clear cache)
2. Wait 5-15 minutes for propagation
3. Check App Catalog for multiple versions - ensure latest is deployed
4. Wait up to 24 hours for full CDN propagation

### Permission Errors

**Possible Causes:**
- User doesn't have required permissions
- App Catalog not accessible
- Site permissions insufficient

**Solutions:**
1. Contact SharePoint administrator
2. Verify user has "Read" permission to view events
3. Verify user has "Contribute" permission to add/edit events
4. Verify user has "Manage Lists" permission if using automatic provisioning

### Recurrence Rule Not Working

**Possible Causes:**
- Incorrect format
- Typo in rule
- Wrong month or weekday

**Solutions:**
1. Verify format: `MONTHLY_BY_NTH_WEEKDAY:MM:WEEKDAY:NTH`
2. Check month is 01-12 (two digits)
3. Check weekday is all caps (MONDAY, TUESDAY, etc.)
4. Check nth is 1-5 (5 = last)
5. Verify EventDate is in the correct month

### Images Not Displaying

**Possible Causes:**
- ImageUrl field empty
- Invalid URL
- Image not accessible

**Solutions:**
1. Verify ImageUrl field has a valid URL
2. Check URL is accessible (try opening in browser)
3. Ensure "Show Images" is enabled in web part settings
4. For default images, they should appear automatically for holidays/birthdays

---

## Getting Help

If you encounter issues not covered in this guide:

1. **Check Browser Console**: Press F12, check Console tab for errors
2. **Contact Your SharePoint Administrator**: For deployment and permission issues
3. **Contact Support**: For the web part, contact: contactus@synozur.com

---

## Quick Reference: Recurrence Patterns

Copy and paste these patterns into the **RecurrenceRule** field:

- **1st Monday**: `MONTHLY_BY_NTH_WEEKDAY:MM:MONDAY:1`
- **2nd Monday**: `MONTHLY_BY_NTH_WEEKDAY:MM:MONDAY:2`
- **3rd Monday**: `MONTHLY_BY_NTH_WEEKDAY:MM:MONDAY:3`
- **4th Monday**: `MONTHLY_BY_NTH_WEEKDAY:MM:MONDAY:4`
- **Last Monday**: `MONTHLY_BY_NTH_WEEKDAY:MM:MONDAY:5`

Replace `MM` with the month number (01-12) and `MONDAY` with the desired weekday.

**Examples:**
- Labor Day (1st Monday of September): `MONTHLY_BY_NTH_WEEKDAY:09:MONDAY:1`
- Thanksgiving (4th Thursday of November): `MONTHLY_BY_NTH_WEEKDAY:11:THURSDAY:4`
- Memorial Day (Last Monday of May): `MONTHLY_BY_NTH_WEEKDAY:05:MONDAY:5`

---

*Last Updated: Version 1.1.1.0*

