# Setup Managers List in SharePoint
# This script creates a SharePoint list to define who is a manager

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$ListName = "Managers"
)

# Import SharePoint PnP PowerShell module
try {
    Import-Module PnP.PowerShell -ErrorAction Stop
    Write-Host "PnP.PowerShell module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Force -AllowClobber
    Import-Module PnP.PowerShell
}

try {
    # Connect to SharePoint site
    Write-Host "Connecting to SharePoint site: $SiteUrl" -ForegroundColor Yellow
    Write-Host "This will open a browser window for authentication..." -ForegroundColor Cyan
    
    # Try different authentication methods
    try {
        # Method 1: Interactive with browser (most compatible)
        Connect-PnPOnline -Url $SiteUrl -Interactive -ForceAuthentication
        Write-Host "Connected successfully using Interactive authentication" -ForegroundColor Green
    }
    catch {
        Write-Host "Interactive authentication failed, trying device login..." -ForegroundColor Yellow
        try {
            # Method 2: Device login as fallback
            Connect-PnPOnline -Url $SiteUrl -DeviceLogin
            Write-Host "Connected successfully using Device Login" -ForegroundColor Green
        }
        catch {
            Write-Host "Device login failed, trying credential prompt..." -ForegroundColor Yellow
            # Method 3: Credential prompt as last resort
            $credential = Get-Credential -Message "Enter your SharePoint credentials"
            Connect-PnPOnline -Url $SiteUrl -Credentials $credential
            Write-Host "Connected successfully using Credentials" -ForegroundColor Green
        }
    }
    
    # Check if list already exists
    $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    
    if ($existingList) {
        Write-Host "List '$ListName' already exists. Do you want to recreate it? (y/n)" -ForegroundColor Yellow
        $response = Read-Host
        if ($response -eq 'y' -or $response -eq 'Y') {
            Remove-PnPList -Identity $ListName -Force
            Write-Host "Existing list removed" -ForegroundColor Green
        } else {
            Write-Host "Script cancelled" -ForegroundColor Red
            return
        }
    }
    
    # Create the Managers list
    Write-Host "Creating '$ListName' list..." -ForegroundColor Yellow
    $list = New-PnPList -Title $ListName -Template GenericList -Description "List to define who is a manager in the organization"
    
    # Add custom columns
    Write-Host "Adding custom columns..." -ForegroundColor Yellow
    
    # Manager Email (Person field)
    Add-PnPField -List $ListName -DisplayName "Manager Email" -InternalName "ManagerEmail" -Type User -Required
    
    # Manager Display Name (Text field)
    Add-PnPField -List $ListName -DisplayName "Manager Display Name" -InternalName "ManagerDisplayName" -Type Text -Required
    
    # Department (Text field)
    Add-PnPField -List $ListName -DisplayName "Department" -InternalName "Department" -Type Text
    
    # Manager Level (Choice field)
    $managerLevels = @("Team Lead", "Department Manager", "Senior Manager", "Director", "VP", "Executive")
    Add-PnPField -List $ListName -DisplayName "Manager Level" -InternalName "ManagerLevel" -Type Choice -Choices $managerLevels
    
    # Is Active (Yes/No field)
    Add-PnPField -List $ListName -DisplayName "Is Active" -InternalName "IsActive" -Type Boolean -Required
    
    # Start Date (Date field)
    Add-PnPField -List $ListName -DisplayName "Start Date" -InternalName "StartDate" -Type DateTime
    
    # End Date (Date field)
    Add-PnPField -List $ListName -DisplayName "End Date" -InternalName "EndDate" -Type DateTime
    
    # Notes (Multi-line text field)
    Add-PnPField -List $ListName -DisplayName "Notes" -InternalName "Notes" -Type Note
    
    # Update the default view to include new columns
    Write-Host "Updating default view..." -ForegroundColor Yellow
    $defaultView = Get-PnPView -List $ListName -Identity "All Items"
    
    # Remove default Title column from view and add our custom columns
    $viewFields = @("ManagerEmail", "ManagerDisplayName", "Department", "ManagerLevel", "IsActive", "StartDate", "EndDate")
    
    Set-PnPView -List $ListName -Identity $defaultView.Id -Fields $viewFields
    
    # Set the Title field to not be required (since we're using custom fields)
    $titleField = Get-PnPField -List $ListName -Identity "Title"
    Set-PnPField -List $ListName -Identity $titleField.Id -Values @{Required=$false}
    
    # Create some sample data (optional)
    Write-Host "Do you want to add sample manager data? (y/n)" -ForegroundColor Yellow
    $addSample = Read-Host
    
    if ($addSample -eq 'y' -or $addSample -eq 'Y') {
        Write-Host "Adding sample data..." -ForegroundColor Yellow
        
        # Note: Replace these with actual user emails from your organization
        $sampleManagers = @(
            @{
                Title = "Sample Manager 1"
                ManagerDisplayName = "John Smith"
                Department = "IT"
                ManagerLevel = "Department Manager"
                IsActive = $true
                StartDate = (Get-Date).AddMonths(-6)
                Notes = "Sample manager entry"
            },
            @{
                Title = "Sample Manager 2"
                ManagerDisplayName = "Jane Doe"
                Department = "HR"
                ManagerLevel = "Senior Manager"
                IsActive = $true
                StartDate = (Get-Date).AddMonths(-12)
                Notes = "Sample manager entry"
            }
        )
        
        foreach ($manager in $sampleManagers) {
            Add-PnPListItem -List $ListName -Values $manager
        }
        
        Write-Host "Sample data added. Please update with real manager information and email addresses." -ForegroundColor Green
    }
    
    # Set list permissions (optional)
    Write-Host "Do you want to set special permissions for this list? (y/n)" -ForegroundColor Yellow
    $setPermissions = Read-Host
    
    if ($setPermissions -eq 'y' -or $setPermissions -eq 'Y') {
        Write-Host "Setting list permissions..." -ForegroundColor Yellow
        
        # Break permission inheritance
        Set-PnPList -Identity $ListName -BreakRoleInheritance -CopyRoleAssignments
        
        Write-Host "Permission inheritance broken. You can now set custom permissions for this list." -ForegroundColor Green
        Write-Host "Consider giving 'Read' access to all users and 'Edit' access only to HR/Admin staff." -ForegroundColor Yellow
    }
    
    $listUrl = "$SiteUrl/Lists/$($ListName)"
    
    Write-Host "SUCCESS!" -ForegroundColor Green
    Write-Host "Managers list created successfully at: $listUrl" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Navigate to the list and add real manager information"
    Write-Host "2. Update the Manager Email field with actual user accounts"
    Write-Host "3. Set appropriate permissions if not done already"
    Write-Host "4. Update your SPFx solution to read from this list"
    Write-Host ""
    Write-Host "List structure created:" -ForegroundColor Cyan
    Write-Host "- Manager Email (Person field) - Required"
    Write-Host "- Manager Display Name (Text) - Required"
    Write-Host "- Department (Text)"
    Write-Host "- Manager Level (Choice)"
    Write-Host "- Is Active (Yes/No) - Required"
    Write-Host "- Start Date (Date)"
    Write-Host "- End Date (Date)"
    Write-Host "- Notes (Multi-line text)"
    
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error details:" -ForegroundColor Red
    Write-Host $_.Exception.ToString()
}
finally {
    # Disconnect from SharePoint
    try {
        $connection = Get-PnPConnection -ErrorAction SilentlyContinue
        if ($connection) {
            Disconnect-PnPOnline
            Write-Host "Disconnected from SharePoint" -ForegroundColor Green
        }
    }
    catch {
        # Ignore disconnect errors
        Write-Host "SharePoint session ended" -ForegroundColor Gray
    }
}
