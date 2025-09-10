# SPFx Adaptive Cards Solution - Setup Guide

## Prerequisites
- Node.js 18.17.1 or higher
- SharePoint Framework development environment
- VS Code (recommended)

## Installation Steps

### 1. Install Dependencies
```bash
cd spfx-adaptivecard-solution
npm install
```

### 2. SharePoint List Setup

Create two SharePoint lists in your target site:

#### Important Messages List
- **List Name**: "Important Messages"
- **Columns**:
  - Title (Single line of text) - Default
  - MessageContent (Multiple lines of text, Enhanced rich text)
  - Priority (Choice: High, Medium, Low)
  - ExpiryDate (Date and Time)
  - TargetAudience (Single line of text)
  - ReadBy (Multiple lines of text, Plain text)

#### Message Read Actions List
- **List Name**: "Message Read Actions"
- **Columns**:
  - Title (Single line of text) - Default
  - MessageId (Number)
  - UserId (Number)
  - UserEmail (Single line of text)
  - UserDisplayName (Single line of text)
  - ReadTimestamp (Date and Time)
  - DeviceInfo (Single line of text)

### 3. Configuration

#### DataService Configuration
In your web part's `onInit()` method, initialize the DataService:

```typescript
protected async onInit(): Promise<void> {
  // Initialize the data service
  const dataService = new DataService();
  dataService.initSP(this.context, 'https://your-power-automate-url'); // Optional Power Automate URL
  
  return super.onInit();
}
```

#### Power Automate Integration (Optional)
1. Create a Power Automate flow with an HTTP trigger
2. Configure the trigger to accept POST requests
3. Use the payload structure defined in `IPowerAutomatePayload`
4. Add the flow URL to the DataService initialization

### 4. Build and Deploy

#### Development
```bash
npm run serve
```

#### Production Build
```bash
npm run build
npm run package-solution
```

## Features

### Adaptive Card Viewer Web Part
- Renders Adaptive Cards from JSON URLs
- Supports interactive elements and actions
- Integration with SharePoint data
- Configurable title and settings

### Dashboard Web Part
- Displays important messages from SharePoint lists
- Read/unread status tracking
- Priority-based styling
- Auto-refresh capabilities
- Mobile-responsive design

### DataService
- SharePoint list operations (CRUD)
- Read tracking functionality
- Power Automate integration
- Error handling and logging

## Usage Examples

### Creating Messages
```typescript
const dataService = new DataService();
await dataService.createMessage({
  Title: "System Maintenance",
  MessageContent: "Scheduled maintenance window...",
  Priority: "High",
  ExpiryDate: new Date("2025-08-15"),
  TargetAudience: "All"
});
```

### Marking Messages as Read
```typescript
await dataService.markMessageAsRead(messageId);
```

### Getting Read Statistics
```typescript
const stats = await dataService.getMessageReadStats(messageId);
console.log(`Total reads: ${stats.totalReads}, Unique readers: ${stats.uniqueReaders}`);
```

## Adaptive Card Examples

### Sample Card JSON
See `src/assets/sample-maintenance-card.json` for a complete example.

### Basic Structure
```json
{
  "type": "AdaptiveCard",
  "version": "1.5",
  "body": [
    {
      "type": "TextBlock",
      "text": "Hello World!",
      "size": "Large",
      "weight": "Bolder"
    }
  ],
  "actions": [
    {
      "type": "Action.Submit",
      "title": "Mark as Read",
      "data": { "action": "markAsRead", "messageId": "123" }
    }
  ]
}
```

## Troubleshooting

### Common Issues

1. **Lists not found**: Ensure SharePoint lists are created with exact names and column types
2. **Permission errors**: Verify the app has appropriate permissions to access SharePoint lists
3. **Build errors**: Check Node.js version compatibility and run `npm install`

### Debugging
- Use browser developer tools to inspect network requests
- Check SharePoint ULS logs for server-side errors
- Enable verbose logging in the DataService

## Customization

### Styling
- Modify SCSS files in the components folders
- Use SharePoint theme tokens for consistent branding
- Responsive design breakpoints are included

### Extending Functionality
- Add new Adaptive Card elements to the renderer
- Implement additional SharePoint list operations
- Create custom Power Automate flows for notifications

## Security Considerations
- Validate all user inputs
- Use SharePoint permissions for access control
- Sanitize HTML content in messages
- Follow SharePoint Framework security guidelines

## Support
For issues and questions:
1. Check the SharePoint Framework documentation
2. Review Adaptive Cards schema reference
3. Consult your organization's SharePoint administrators
