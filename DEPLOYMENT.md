# Holidays & Birthdays Web Part - Deployment Guide

## Quick Start

### 1. Build the Solution

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` file will be created in the `solution` folder.

### 2. Deploy to App Catalog

#### For Tenant-Wide Deployment (Recommended for Production)

1. Navigate to **SharePoint Admin Center** → **More features** → **Apps** → **App Catalog**
   - Or go directly to: `https://[tenant]-admin.sharepoint.com/_layouts/15/tenantAppCatalog.aspx`
2. Click **Distributed apps** or **Apps for SharePoint**
3. Click **Upload** or drag-and-drop the new `.sppkg` file
4. **Important**: When prompted, check **"Replace the existing file if it exists"**
5. Click **Deploy** when prompted
6. Check **"Make this solution available to all sites in the organization"** for tenant-wide deployment
7. Click **OK** to deploy

#### For Site Collection App Catalog

1. Navigate to your SharePoint site
2. Go to **Site Contents** → **Site App Catalog** (or Site Collection App Catalog)
3. Click **Upload** or drag-and-drop the `.sppkg` file
4. **Important**: When prompted, check **"Replace the existing file if it exists"**
5. Click **Deploy** when prompted
6. Optionally check "Make this solution available to all sites in the organization" if deploying tenant-wide

## Updating an Existing Deployment (Best Practices)

### Pre-Deployment Checklist

- [ ] **Version Number**: Ensure version is incremented in `package.json` and `package-solution.json`
- [ ] **Test Locally**: Test the web part locally using `gulp serve` before packaging
- [ ] **Build Clean**: Run `gulp clean` before building to ensure fresh build
- [ ] **Documentation**: Update changelog/README with new features/changes
- [ ] **Backup**: Consider backing up existing web part instances (export page configurations if needed)

### Deployment Steps for Updates

1. **Build the Updated Package**:
   ```bash
   gulp clean
   gulp bundle --ship
   gulp package-solution --ship
   ```

2. **Upload to App Catalog**:
   - Navigate to App Catalog (Tenant or Site Collection)
   - Upload the new `.sppkg` file
   - **Check "Replace the existing file if it exists"** (critical for updates)
   - Click **Deploy**

3. **Post-Deployment**:
   - **Clear Browser Cache**: Users may need to clear browser cache or do hard refresh (Ctrl+F5)
   - **Wait for Propagation**: Changes may take 5-15 minutes to propagate across all sites
   - **Test on Multiple Sites**: Verify the update works on different sites
   - **Check Browser Console**: Look for any JavaScript errors after update

### Important Notes for Updates

- **No Downtime**: SPFx updates are deployed without downtime - existing instances continue working
- **Automatic Updates**: Once deployed, all instances of the web part automatically use the new version
- **Property Pane Changes**: If you changed property pane settings, existing instances may need to be reconfigured
- **Breaking Changes**: If you made breaking changes to properties, existing web parts may show errors until reconfigured
- **CDN Caching**: SharePoint CDN caches assets - updates may take time to propagate globally

### Rollback Procedure

If you need to rollback to a previous version:

1. Build the previous version from source control
2. Upload and deploy the previous `.sppkg` file
3. Check "Replace the existing file if it exists"
4. Deploy
5. Users may need to clear browser cache

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
- **EventDate**: 2026-01-01 (any year works)
- **EventType**: Holiday
- **IsAnnualRecurrence**: Yes
- **RecurrenceRule**: (leave empty)

### Moving Holiday (Labor Day - 1st Monday of September)

- **Title**: Labor Day
- **EventDate**: 2026-09-01 (any date in September)
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

### Deployment Issues

#### "File Already Exists" Error
- **Solution**: Check "Replace the existing file if it exists" when uploading
- If option doesn't appear, delete the old package first, then upload new one

#### Web Part Not Updating After Deployment
- **Clear Browser Cache**: Hard refresh (Ctrl+F5) or clear browser cache
- **Wait for Propagation**: Can take 5-15 minutes for changes to propagate
- **Check Version**: Verify the deployed version matches your build output
- **CDN Cache**: SharePoint CDN may cache assets - wait up to 24 hours for full propagation

#### "App Not Available" After Update
- Verify deployment completed successfully in App Catalog
- Check if solution is marked as "Available" in App Catalog
- Ensure "Make this solution available to all sites" is checked (if tenant-wide)
- Check browser console for JavaScript errors

### Runtime Issues

#### "List not found" Error

- Verify list name matches exactly (case-sensitive)
- Enable "Allow List Provisioning" in web part settings
- Check user permissions to create lists

#### Events Not Appearing

- Verify items exist in the list
- Check EventDate format (YYYY-MM-DD)
- Ensure IsAnnualRecurrence is Yes for recurring events
- Verify RecurrenceRule format if using moving holidays

#### Permission Errors

- Contact SharePoint administrator
- Ensure web part is deployed from accessible app catalog
- Verify user has appropriate site permissions

#### Web Part Shows Old Version
- Clear browser cache (Ctrl+F5)
- Check if multiple versions exist in App Catalog
- Verify the correct version is deployed
- Wait for CDN propagation (can take up to 24 hours)

## Development

For local development:

```bash
npm install
gulp serve
```

Update `config/serve.json` with your tenant URL for testing.

