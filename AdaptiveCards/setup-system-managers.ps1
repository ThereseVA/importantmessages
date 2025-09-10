# 👑 System Managers List Setup Script
# This script creates the required SharePoint list for managing user roles and permissions

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateSampleData,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

Write-Host "👑 Setting up System Managers list for user role management..." -ForegroundColor Green
Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow

# Connect to SharePoint
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive
    Write-Host "✅ Connected to SharePoint successfully" -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to connect to SharePoint: $($_.Exception.Message)"
    exit 1
}

# Create System Managers List
Write-Host "`n👑 Creating 'SystemManagers' list..." -ForegroundColor Cyan
try {
    $managersList = Get-PnPList -Identity "SystemManagers" -ErrorAction SilentlyContinue
    if ($managersList -and -not $Force) {
        Write-Host "⚠️ 'SystemManagers' list already exists. Use -Force to recreate." -ForegroundColor Yellow
    } else {
        if ($managersList -and $Force) {
            Write-Host "🗑️ Removing existing 'SystemManagers' list..." -ForegroundColor Yellow
            Remove-PnPList -Identity "SystemManagers" -Force
            Start-Sleep -Seconds 2
        }
        
        # Create new list
        New-PnPList -Title "SystemManagers" -Template GenericList -OnQuickLaunch:$false
        Write-Host "✅ Created 'SystemManagers' list" -ForegroundColor Green
        
        # Add custom columns
        Write-Host "📝 Adding custom columns..." -ForegroundColor Cyan
        
        # UserEmail - Single line of text (required)
        Add-PnPField -List "SystemManagers" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -Required -AddToDefaultView
        
        # UserDisplayName - Single line of text (required)
        Add-PnPField -List "SystemManagers" -DisplayName "UserDisplayName" -InternalName "UserDisplayName" -Type Text -Required -AddToDefaultView
        
        # Role - Choice field (required)
        Add-PnPField -List "SystemManagers" -DisplayName "Role" -InternalName "Role" -Type Choice -Choices @("Manager","Admin","SuperAdmin") -DefaultValue "Manager" -Required -AddToDefaultView
        
        # Department - Single line of text (optional)
        Add-PnPField -List "SystemManagers" -DisplayName "Department" -InternalName "Department" -Type Text -AddToDefaultView
        
        # IsActive - Boolean field (required, default true)
        Add-PnPField -List "SystemManagers" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -DefaultValue $true -Required -AddToDefaultView
        
        Write-Host "✅ Added custom columns to 'SystemManagers' list" -ForegroundColor Green
        
        # Set up list permissions and settings
        Write-Host "🔧 Configuring list settings..." -ForegroundColor Cyan
        
        # Hide from quick launch
        Set-PnPList -Identity "SystemManagers" -Hidden:$false -OnQuickLaunch:$false
        
        # Set list description
        Set-PnPList -Identity "SystemManagers" -Description "System configuration list for managing user roles and permissions. Controls who has Manager, Admin, or SuperAdmin access to the Adaptive Cards solution."
        
        Write-Host "✅ List configuration completed" -ForegroundColor Green
    }
} catch {
    Write-Error "❌ Failed to create 'SystemManagers' list: $($_.Exception.Message)"
}

# Create sample data if requested
if ($CreateSampleData) {
    Write-Host "`n👤 Creating sample manager data..." -ForegroundColor Cyan
    try {
        # Sample managers - Update these emails to match your organization
        $sampleManagers = @(
            @{
                Title = "Super Administrator"
                UserEmail = "admin@gustafkliniken.sharepoint.com"
                UserDisplayName = "System Administrator"
                Role = "SuperAdmin"
                Department = "IT"
                IsActive = $true
            },
            @{
                Title = "Manager - Therese Almesjo"
                UserEmail = "therese.almesjo@gustafkliniken.sharepoint.com"
                UserDisplayName = "Therese Almesjo"
                Role = "Admin"
                Department = "Management"
                IsActive = $true
            },
            @{
                Title = "Department Manager"
                UserEmail = "manager@gustafkliniken.sharepoint.com"
                UserDisplayName = "Department Manager"
                Role = "Manager"
                Department = "Operations"
                IsActive = $true
            }
        )
        
        foreach ($manager in $sampleManagers) {
            try {
                # Check if manager already exists
                $existingManager = Get-PnPListItem -List "SystemManagers" -Query "<View><Query><Where><Eq><FieldRef Name='UserEmail'/><Value Type='Text'>$($manager.UserEmail)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                
                if ($existingManager) {
                    Write-Host "⚠️ Manager already exists: $($manager.UserEmail)" -ForegroundColor Yellow
                } else {
                    Add-PnPListItem -List "SystemManagers" -Values $manager
                    Write-Host "✅ Added manager: $($manager.UserDisplayName) ($($manager.Role))" -ForegroundColor Green
                }
            } catch {
                Write-Host "⚠️ Could not add manager $($manager.UserEmail): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✅ Sample data creation completed" -ForegroundColor Green
    } catch {
        Write-Error "❌ Failed to create sample data: $($_.Exception.Message)"
    }
}

# Display summary and next steps
Write-Host "`n📋 SETUP SUMMARY" -ForegroundColor Green
Write-Host "=================" -ForegroundColor Green
Write-Host "✅ SystemManagers list created with columns:" -ForegroundColor White
Write-Host "   • UserEmail (Text, Required)" -ForegroundColor Gray
Write-Host "   • UserDisplayName (Text, Required)" -ForegroundColor Gray
Write-Host "   • Role (Choice: Manager/Admin/SuperAdmin, Required)" -ForegroundColor Gray
Write-Host "   • Department (Text, Optional)" -ForegroundColor Gray
Write-Host "   • IsActive (Boolean, Required, Default: True)" -ForegroundColor Gray

Write-Host "`n🎯 NEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. 👤 Add your managers to the 'SystemManagers' list" -ForegroundColor White
Write-Host "2. 🔄 Deploy the updated SPFx solution with UserRoleService" -ForegroundColor White
Write-Host "3. ✅ Test manager detection with different user accounts" -ForegroundColor White
Write-Host "4. 🚀 The system will automatically check this list first for manager permissions" -ForegroundColor White

Write-Host "`n💡 MANAGEMENT TIPS:" -ForegroundColor Cyan
Write-Host "• Use 'SuperAdmin' role for full system access" -ForegroundColor White
Write-Host "• Use 'Admin' role for department-level management" -ForegroundColor White
Write-Host "• Use 'Manager' role for basic manager permissions" -ForegroundColor White
Write-Host "• Set IsActive to False to temporarily disable a manager without deleting" -ForegroundColor White
Write-Host "• The system checks: SharePoint List → Groups → Hardcoded fallback" -ForegroundColor White

Write-Host "`n🔗 List URL: $SiteUrl/Lists/SystemManagers" -ForegroundColor Green

# Disconnect from SharePoint
Disconnect-PnPOnline
Write-Host "`n✅ Setup completed successfully!" -ForegroundColor Green
