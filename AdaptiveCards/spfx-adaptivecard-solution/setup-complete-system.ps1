# COMPREHENSIVE SharePoint Lists Setup Script
# This script creates ALL required lists for the complete Teams/Outlook/SharePoint solution

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateSampleData,
    
    [Parameter(Mandatory=$false)]
    [switch]$ForceRecreate
)

Write-Host "🚀 COMPREHENSIVE SETUP - Teams/Outlook/SharePoint Solution" -ForegroundColor Green
Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow
Write-Host "Force Recreate: $ForceRecreate" -ForegroundColor Yellow

# Connect to SharePoint
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive
    Write-Host "✅ Connected to SharePoint successfully" -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to connect to SharePoint: $($_.Exception.Message)"
    exit 1
}

# Function to create or recreate list
function New-OrRecreateList {
    param($ListName, $Template = "GenericList")
    
    $existingList = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($existingList) {
        if ($ForceRecreate) {
            Write-Host "🗑️ Deleting existing '$ListName' list..." -ForegroundColor Yellow
            Remove-PnPList -Identity $ListName -Force
            Start-Sleep -Seconds 2
        } else {
            Write-Host "⚠️ '$ListName' list already exists (use -ForceRecreate to recreate)" -ForegroundColor Yellow
            return $existingList
        }
    }
    
    Write-Host "📋 Creating '$ListName' list..." -ForegroundColor Cyan
    $newList = New-PnPList -Title $ListName -Template $Template
    Write-Host "✅ Created '$ListName' list" -ForegroundColor Green
    return $newList
}

# 1. Create Important Messages List (MAIN MESSAGE STORAGE)
Write-Host "`n📋 SETTING UP IMPORTANT MESSAGES LIST" -ForegroundColor Magenta
$messagesList = New-OrRecreateList -ListName "Important Messages"

if ($messagesList) {
    try {
        # Add custom columns with proper types and choices
        Add-PnPField -List "Important Messages" -DisplayName "MessageContent" -InternalName "MessageContent" -Type Note -AddToDefaultView
        Add-PnPField -List "Important Messages" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices @("High","Medium","Low") -DefaultValue "Medium" -AddToDefaultView
        Add-PnPField -List "Important Messages" -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime -AddToDefaultView
        
        # TargetAudience with proper choices matching Channel Groups
        Add-PnPField -List "Important Messages" -DisplayName "TargetAudience" -InternalName "TargetAudience" -Type Choice -Choices @("All Users","Alla Medarbetare","Läkare","Sjuksköterskor","Administration","Reception","Ledning","IT Support") -DefaultValue "All Users" -AddToDefaultView
        
        # Source tracking
        Add-PnPField -List "Important Messages" -DisplayName "Source" -InternalName "Source" -Type Choice -Choices @("SharePoint","Teams","Outlook","Manual") -DefaultValue "SharePoint" -AddToDefaultView
        
        # Read tracking fields
        Add-PnPField -List "Important Messages" -DisplayName "ReadBy" -InternalName "ReadBy" -Type Note
        Add-PnPField -List "Important Messages" -DisplayName "TotalReads" -InternalName "TotalReads" -Type Number -DefaultValue "0"
        Add-PnPField -List "Important Messages" -DisplayName "UniqueReaders" -InternalName "UniqueReaders" -Type Number -DefaultValue "0"
        
        Write-Host "✅ Added custom columns to 'Important Messages' list" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Some columns may already exist: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# 2. Create MessageReadConfirmations List (READ TRACKING)
Write-Host "`n📊 SETTING UP MESSAGE READ CONFIRMATIONS LIST" -ForegroundColor Magenta
$readActionsList = New-OrRecreateList -ListName "MessageReadConfirmations"

if ($readActionsList) {
    try {
        # Add read tracking columns
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "MessageId" -InternalName "MessageId" -Type Number -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "UserId" -InternalName "UserId" -Type Number -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "UserDisplayName" -InternalName "UserDisplayName" -Type Text -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "ReadTimestamp" -InternalName "ReadTimestamp" -Type DateTime -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "DeviceInfo" -InternalName "DeviceInfo" -Type Note
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "Platform" -InternalName "Platform" -Type Choice -Choices @("SharePoint","Teams","Outlook","Mobile") -DefaultValue "SharePoint"
        
        Write-Host "✅ Added custom columns to 'MessageReadConfirmations' list" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Some columns may already exist: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# 3. Create Channel Groups List (TARGET AUDIENCE MANAGEMENT)
Write-Host "`n🎯 SETTING UP CHANNEL GROUPS LIST" -ForegroundColor Magenta
$channelGroupsList = New-OrRecreateList -ListName "Channel Groups"

if ($channelGroupsList) {
    try {
        # Add group management columns
        Add-PnPField -List "Channel Groups" -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView
        Add-PnPField -List "Channel Groups" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -DefaultValue $true -AddToDefaultView
        Add-PnPField -List "Channel Groups" -DisplayName "SortOrder" -InternalName "SortOrder" -Type Number -DefaultValue "0" -AddToDefaultView
        Add-PnPField -List "Channel Groups" -DisplayName "TeamsChannelId" -InternalName "TeamsChannelId" -Type Text
        Add-PnPField -List "Channel Groups" -DisplayName "EmailDistributionList" -InternalName "EmailDistributionList" -Type Note
        Add-PnPField -List "Channel Groups" -DisplayName "UserCount" -InternalName "UserCount" -Type Number -DefaultValue "0"
        
        Write-Host "✅ Added custom columns to 'Channel Groups' list" -ForegroundColor Green
        
        # Add default channel groups
        Write-Host "🎯 Adding default channel groups..." -ForegroundColor Cyan
        
        $defaultGroups = @(
            @{ Title="Alla Medarbetare"; Description="Alla anställda på Gustaf Kliniken"; SortOrder=1; IsActive=$true },
            @{ Title="Läkare"; Description="Alla läkare och medicinskt ansvariga"; SortOrder=2; IsActive=$true },
            @{ Title="Sjuksköterskor"; Description="Sjuksköterskor och omvårdnadspersonal"; SortOrder=3; IsActive=$true },
            @{ Title="Administration"; Description="Administrativ personal och ekonomi"; SortOrder=4; IsActive=$true },
            @{ Title="Reception"; Description="Reception och patientmottagning"; SortOrder=5; IsActive=$true },
            @{ Title="Ledning"; Description="Chefer och verksamhetsledning"; SortOrder=6; IsActive=$true },
            @{ Title="IT Support"; Description="IT-personal och teknisk support"; SortOrder=7; IsActive=$true }
        )
        
        foreach ($group in $defaultGroups) {
            try {
                Add-PnPListItem -List "Channel Groups" -Values $group
                Write-Host "  ✅ Added group: $($group.Title)" -ForegroundColor Green
            } catch {
                Write-Host "  ⚠️ Group may already exist: $($group.Title)" -ForegroundColor Yellow
            }
        }
        
    } catch {
        Write-Host "⚠️ Some columns may already exist: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# 4. Create TeamsDistributionLogs List (TEAMS INTEGRATION LOGS)
Write-Host "`n📤 SETTING UP TEAMS DISTRIBUTION LOGS LIST" -ForegroundColor Magenta
$distributionLogsList = New-OrRecreateList -ListName "TeamsDistributionLogs"

if ($distributionLogsList) {
    try {
        # Add distribution tracking columns
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "MessageId" -InternalName "MessageId" -Type Number -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ChannelUrl" -InternalName "ChannelUrl" -Type URL -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ChannelName" -InternalName "ChannelName" -Type Text -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "DistributionStatus" -InternalName "DistributionStatus" -Type Choice -Choices @("Success","Failed","Pending","Retrying") -DefaultValue "Pending" -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "DistributionTimestamp" -InternalName "DistributionTimestamp" -Type DateTime -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ErrorMessage" -InternalName "ErrorMessage" -Type Note
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ResponseData" -InternalName "ResponseData" -Type Note
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "RetryCount" -InternalName "RetryCount" -Type Number -DefaultValue "0"
        
        Write-Host "✅ Added custom columns to 'TeamsDistributionLogs' list" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Some columns may already exist: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# 5. Create sample data if requested
if ($CreateSampleData) {
    Write-Host "`n🎯 CREATING SAMPLE DATA" -ForegroundColor Magenta
    
    try {
        # Get current date and future dates
        $now = Get-Date
        $tomorrow = $now.AddDays(1)
        $nextWeek = $now.AddDays(7)
        
        # Sample messages
        $sampleMessages = @(
            @{
                Title = "Viktigt meddelande: Nya rutiner för patientmottagning"
                MessageContent = "Från och med måndag den 5 augusti implementerar vi nya rutiner för patientmottagning. Alla medarbetare ska läsa och bekräfta att de har tagit del av informationen."
                Priority = "High"
                TargetAudience = "Alla Medarbetare"
                ExpiryDate = $nextWeek.ToString("yyyy-MM-ddTHH:mm:ssZ")
                Source = "SharePoint"
            },
            @{
                Title = "Medicinteknik: Uppdatering av system"
                MessageContent = "Våra medicinska system kommer att uppdateras under helgen. Var vänlig kontakta IT-supporten vid eventuella problem."
                Priority = "Medium"
                TargetAudience = "Läkare"
                ExpiryDate = $tomorrow.ToString("yyyy-MM-ddTHH:mm:ssZ")
                Source = "Teams"
            },
            @{
                Title = "Schema: Nya arbetstider från nästa vecka"
                MessageContent = "På grund av sommarsemestrar justeras arbetstiderna tillfälligt. Kontrollera era scheman i personalplaneringssystemet."
                Priority = "Medium"
                TargetAudience = "Sjuksköterskor"
                ExpiryDate = $nextWeek.ToString("yyyy-MM-ddTHH:mm:ssZ")
                Source = "Outlook"
            }
        )
        
        foreach ($message in $sampleMessages) {
            try {
                $newItem = Add-PnPListItem -List "Important Messages" -Values $message
                Write-Host "  ✅ Created sample message: $($message.Title)" -ForegroundColor Green
            } catch {
                Write-Host "  ⚠️ Failed to create sample message: $($message.Title)" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✅ Sample data created successfully" -ForegroundColor Green
        
    } catch {
        Write-Error "❌ Failed to create sample data: $($_.Exception.Message)"
    }
}

Write-Host "`n🎉 COMPREHENSIVE SETUP COMPLETED!" -ForegroundColor Green
Write-Host "`n📋 CREATED LISTS:" -ForegroundColor Cyan
Write-Host "  ✅ Important Messages (main message storage)" -ForegroundColor Green
Write-Host "  ✅ MessageReadConfirmations (read tracking)" -ForegroundColor Green
Write-Host "  ✅ Channel Groups (target audience management)" -ForegroundColor Green
Write-Host "  ✅ TeamsDistributionLogs (Teams integration logs)" -ForegroundColor Green

Write-Host "`n🔧 NEXT STEPS:" -ForegroundColor Yellow
Write-Host "  1. Deploy the updated SPFx package" -ForegroundColor White
Write-Host "  2. Configure Power Automate flows with correct list names" -ForegroundColor White
Write-Host "  3. Install Teams app with updated manifest" -ForegroundColor White
Write-Host "  4. Test the complete integration" -ForegroundColor White

Write-Host "`n🎯 The system is now ready for Teams/Outlook/SharePoint integration!" -ForegroundColor Green
