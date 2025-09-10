# SharePoint Managers List Setup Guide

This guide explains how to set up and use the SharePoint Managers list to define who is a manager in your organization.

## Overview

The Managers list is a SharePoint list that defines which users have manager privileges in your important messages system. This replaces hardcoded manager checks with a flexible, configurable solution.

## Setup Steps

### 1. Run the Setup Script

Execute the PowerShell script to create the Managers list:

```powershell
.\setup-managers-list.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/yoursite"
```

### 2. List Structure

The script creates a list with the following columns:

| Column Name | Type | Required | Description |
|-------------|------|----------|-------------|
| Manager Email | Person | Yes | The user's email/account (Person field) |
| Manager Display Name | Text | Yes | Display name of the manager |
| Department | Text | No | Department the manager belongs to |
| Manager Level | Choice | No | Level: Team Lead, Department Manager, Senior Manager, Director, VP, Executive |
| Is Active | Yes/No | Yes | Whether the manager role is currently active |
| Start Date | Date | No | When the manager role started |
| End Date | Date | No | When the manager role ended (for inactive managers) |
| Notes | Multi-line Text | No | Additional notes about the manager role |

### 3. Add Manager Data

After running the script:

1. Navigate to the Managers list in SharePoint
2. Add entries for each manager in your organization
3. Fill in the required fields:
   - **Manager Email**: Select the user from the people picker
   - **Manager Display Name**: Enter their display name
   - **Is Active**: Set to "Yes" for current managers

### 4. Set Permissions (Recommended)

Consider setting these permissions on the Managers list:

- **All Users**: Read access (to check manager status)
- **HR/Admin Staff**: Edit access (to manage the list)
- **System Administrators**: Full Control

## Usage in Code

### Check if Current User is Manager

```typescript
import { enhancedDataService } from '../services/EnhancedDataService';

// Check if current user is a manager
const isManager = await enhancedDataService.isCurrentUserManager();

if (isManager) {
  // Show manager-specific content
  console.log('User is a manager');
}
```

### Check if Specific User is Manager

```typescript
// Check if a specific user is a manager
const isUserManager = await enhancedDataService.isUserManager('john.doe@company.com');
```

### Get Manager Details

```typescript
// Get manager details for current user
const managerDetails = await enhancedDataService.getCurrentUserManagerDetails();

if (managerDetails) {
  console.log(`Manager Level: ${managerDetails.ManagerLevel}`);
  console.log(`Department: ${managerDetails.Department}`);
}
```

### Get All Managers

```typescript
// Get list of all active managers
const allManagers = await enhancedDataService.getAllManagers();

allManagers.forEach(manager => {
  console.log(`${manager.ManagerDisplayName} - ${manager.Department}`);
});
```

## Integration with Web Parts

The Enhanced Data Service automatically checks the Managers list when determining user permissions. This affects:

1. **Manager Dashboard Access**: Only users in the Managers list with `IsActive = true` can access the Manager Dashboard
2. **Message Creation**: Manager-only message creation features
3. **Administrative Functions**: Manager-specific administrative capabilities

## Best Practices

### 1. Regular Maintenance
- Review the list quarterly to ensure accuracy
- Set end dates for managers who are no longer in those roles
- Use the "Is Active" field to temporarily disable manager access

### 2. Department Organization
- Use consistent department names
- Consider creating a separate Department list for standardization

### 3. Manager Levels
- Use the Manager Level field to implement different permission levels
- Consider different dashboard views based on manager level

### 4. Backup and Audit
- Regularly export the list data for backup
- Keep audit trail of changes using SharePoint's version history

## Troubleshooting

### Common Issues

1. **Manager Dashboard Not Accessible**
   - Verify the user is listed in the Managers list
   - Check that `IsActive` is set to "Yes"
   - Ensure the user's email matches exactly

2. **Permission Errors**
   - Verify the user has read access to the Managers list
   - Check that the list exists and is accessible

3. **Script Execution Errors**
   - Ensure PnP PowerShell module is installed
   - Verify you have permissions to create lists on the site
   - Check that you're connected to the correct SharePoint site

### Debug Steps

1. **Check List Exists**:
   ```typescript
   try {
     const managers = await enhancedDataService.getAllManagers();
     console.log('Managers list accessible:', managers.length > 0);
   } catch (error) {
     console.error('Managers list error:', error);
   }
   ```

2. **Verify User in List**:
   ```typescript
   const userEmail = 'user@company.com';
   const isManager = await enhancedDataService.isUserManager(userEmail);
   console.log(`${userEmail} is manager:`, isManager);
   ```

## Migration from Hardcoded Managers

If you previously had hardcoded manager checks:

1. **Identify Current Managers**: List all users who currently have manager access
2. **Create List Entries**: Add them to the Managers list using the setup script
3. **Update Code**: Replace hardcoded checks with list-based checks
4. **Test**: Verify all managers can still access their features
5. **Remove Hardcoded Logic**: Clean up old permission checking code

## Example Data

Here's sample data structure for the Managers list:

| Manager Email | Manager Display Name | Department | Manager Level | Is Active | Start Date |
|---------------|---------------------|------------|---------------|-----------|------------|
| john.smith@company.com | John Smith | IT | Department Manager | Yes | 2024-01-15 |
| jane.doe@company.com | Jane Doe | HR | Senior Manager | Yes | 2023-06-01 |
| mike.wilson@company.com | Mike Wilson | Sales | Team Lead | Yes | 2024-03-01 |

## Security Considerations

1. **Access Control**: Limit who can edit the Managers list
2. **Regular Reviews**: Periodically review and validate manager assignments
3. **Audit Trail**: Use SharePoint's built-in versioning and audit features
4. **Principle of Least Privilege**: Only grant manager access when necessary

## Support

For issues with the Managers list setup or usage:

1. Check SharePoint permissions
2. Verify the list structure matches the expected schema
3. Review console logs for detailed error messages
4. Ensure the Enhanced Data Service is properly initialized
