# Revlon Project Tracker - SPFx Web Part

A SharePoint Framework web part for tracking and managing Revlon projects.

<img width="1238" height="525" alt="image" src="https://github.com/user-attachments/assets/9679606e-7a6d-4ad6-90c1-be430875b73a" />

<img width="906" height="389" alt="image" src="https://github.com/user-attachments/assets/d9a8cc22-bf49-4c39-a67b-8bdfb35d1265" />

<img width="1119" height="304" alt="image" src="https://github.com/user-attachments/assets/d46142b5-fb2d-4e4a-8e13-48d0a45ab578" />

## Prerequisites

- Node.js 18.17.1 or higher
- npm or yarn package manager
- SharePoint Online tenant access

## Installation

1. Install dependencies:
   ```bash
   npm install
   ```

## Building the Project

### Development Build
```bash
gulp build
```

### Production Build
```bash
gulp bundle --ship
gulp package-solution --ship
```

The package file will be created at: `sharepoint/solution/revlon-project-tracker.sppkg`

## Deployment

### Step 1: Upload to App Catalog

1. Navigate to your SharePoint App Catalog:
   - Go to: `https://yourtenant.sharepoint.com/sites/appcatalog`
   - Or via SharePoint Admin Center > Apps > App Catalog

2. Upload the package:
   - Click **Apps for SharePoint**
   - Click **New** or **Upload**
   - Select: `sharepoint/solution/revlon-project-tracker.sppkg`
   - Check **"Make this solution available to all sites in the organization"**
   - Click **Deploy**

### Step 2: Add Web Part to Page

1. Navigate to any SharePoint site page
2. Click **Edit** (top right)
3. Click **+** to add a web part
4. Search for **"Revlon Project Tracker"**
5. Add it to the page

## SharePoint List Setup

The web part requires a SharePoint list named **"RevlonProjects"** in the same site where the web part is used.

### Create the List

1. Go to your SharePoint site
2. Click **Settings** (gear icon) > **Site contents**
3. Click **New** > **List**
4. Choose **Blank list**
5. Name it: **RevlonProjects** (exact spelling, case-sensitive)
6. Click **Create**

### Add Required Columns

After creating the list, add these columns:

1. **Title** - Already exists (default column)
2. **ProjectStatus** - Type: Single line of text
3. **ProjectManager** - Type: Single line of text
4. **StartDate** - Type: Date and Time
5. **EndDate** - Type: Date and Time

**Important:** Column names must match exactly (case-sensitive, no spaces).

### Verify Permissions

Ensure users have **Contribute** or **Full Control** permission on the list:

1. Go to the **RevlonProjects** list
2. Click **Settings** > **List settings**
3. Under **Permissions and Management**, click **Permissions for this list**
4. Verify users have **Contribute** or **Full Control** permission

## Features

- **Project List View**: Display all projects in a grid
- **Add New Projects**: Create new project entries
- **Mock Data Support**: Works offline with sample data
- **Error Handling**: Shows helpful messages if list is missing
- **Modern UI**: Uses Fluent UI components

## Troubleshooting

### Web Part Not Appearing

- Wait 2-3 minutes after deployment
- Clear browser cache (Ctrl+F5 or Cmd+Shift+R)
- Verify package is deployed in App Catalog

### "List Not Found" Error

- Verify list name is exactly: **RevlonProjects**
- Ensure list is in the same site as the web part
- Check list exists in Site Contents

### "Failed to Add Project" Error

- Verify column names match exactly (case-sensitive)
- Check you have "Add Items" permission on the list
- Open browser console (F12) for detailed error messages

### List in Different Site

The list must be in the **same site** where the web part is displayed. If you see a 404 error, check:
- The site URL where the web part is running
- Create the list in that exact site

## Development

### Local Testing

1. Start development server:
   ```bash
   gulp serve
   ```

2. Navigate to SharePoint hosted workbench:
   ```
   https://yourtenant.sharepoint.com/_layouts/workbench.aspx
   ```

3. Add the web part to test

### Clean Build

```bash
gulp clean
gulp build
```

## Project Structure

```
src/
├── services/              # Service layer (data access)
│   ├── IProjectService.ts
│   ├── ProjectService.ts
│   └── pnpjsConfig.ts
└── webparts/
    └── revlonProjectTracker/
        ├── components/
        │   └── RevlonProjectTracker.tsx
        └── RevlonProjectTrackerWebPart.ts
```

## Version

Current version: **1.0.2** (displayed in web part footer)
