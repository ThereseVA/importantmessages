# SharePoint Dashboard Lists Setup Script
# This script creates the required lists for the Dashboard web part

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateSampleData
)

Write-Host "🚀 Setting up SharePoint lists for Dashboard web part..." -ForegroundColor Green
Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow

# Connect to SharePoint
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive
    Write-Host "✅ Connected to SharePoint successfully" -ForegroundColor Green
} catch {
    Write-Error "❌ Failed to connect to SharePoint: $($_.Exception.Message)"
    exit 1
}

# Create Important Messages List
Write-Host "`n📋 Creating 'Important Messages' list..." -ForegroundColor Cyan
try {
    $messagesList = Get-PnPList -Identity "Important Messages" -ErrorAction SilentlyContinue
    if ($messagesList) {
        Write-Host "⚠️ 'Important Messages' list already exists" -ForegroundColor Yellow
    } else {
        $messagesList = New-PnPList -Title "Important Messages" -Template GenericList
        Write-Host "✅ Created 'Important Messages' list" -ForegroundColor Green
        
        # Add custom columns
        Add-PnPField -List "Important Messages" -DisplayName "MessageContent" -InternalName "MessageContent" -Type Note -AddToDefaultView
        Add-PnPField -List "Important Messages" -DisplayName "Priority" -InternalName "Priority" -Type Choice -Choices @("High","Medium","Low") -DefaultValue "Medium" -AddToDefaultView
        Add-PnPField -List "Important Messages" -DisplayName "ExpiryDate" -InternalName "ExpiryDate" -Type DateTime -AddToDefaultView
        Add-PnPField -List "Important Messages" -DisplayName "TargetAudience" -InternalName "TargetAudience" -Type Choice -Choices @("All Users","Alla Medarbetare","Läkare","Sjuksköterskor","Administration","Reception","Ledning") -DefaultValue "All Users" -AddToDefaultView
        Add-PnPField -List "Important Messages" -DisplayName "ReadBy" -InternalName "ReadBy" -Type Note
        
        Write-Host "✅ Added custom columns to 'Important Messages' list" -ForegroundColor Green
    }
} catch {
    Write-Error "❌ Failed to create 'Important Messages' list: $($_.Exception.Message)"
}

# Create MessageReadConfirmations List
Write-Host "`n📊 Creating 'MessageReadConfirmations' list..." -ForegroundColor Cyan
try {
    $readActionsList = Get-PnPList -Identity "MessageReadConfirmations" -ErrorAction SilentlyContinue
    if ($readActionsList) {
        Write-Host "⚠️ 'MessageReadConfirmations' list already exists" -ForegroundColor Yellow
    } else {
        $readActionsList = New-PnPList -Title "MessageReadConfirmations" -Template GenericList
        Write-Host "✅ Created 'MessageReadConfirmations' list" -ForegroundColor Green
        
        # Add custom columns
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "MessageId" -InternalName "MessageId" -Type Number -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "UserId" -InternalName "UserId" -Type Number -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "UserEmail" -InternalName "UserEmail" -Type Text -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "UserDisplayName" -InternalName "UserDisplayName" -Type Text -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "ReadTimestamp" -InternalName "ReadTimestamp" -Type DateTime -AddToDefaultView
        Add-PnPField -List "MessageReadConfirmations" -DisplayName "DeviceInfo" -InternalName "DeviceInfo" -Type Note
        
        Write-Host "✅ Added custom columns to 'MessageReadConfirmations' list" -ForegroundColor Green
    }
} catch {
    Write-Error "❌ Failed to create 'MessageReadConfirmations' list: $($_.Exception.Message)"
}

# Create TeamsDistributionLogs List
Write-Host "`n📤 Creating 'TeamsDistributionLogs' list..." -ForegroundColor Cyan
try {
    $distributionLogsList = Get-PnPList -Identity "TeamsDistributionLogs" -ErrorAction SilentlyContinue
    if ($distributionLogsList) {
        Write-Host "⚠️ 'TeamsDistributionLogs' list already exists" -ForegroundColor Yellow
    } else {
        $distributionLogsList = New-PnPList -Title "TeamsDistributionLogs" -Template GenericList
        Write-Host "✅ Created 'TeamsDistributionLogs' list" -ForegroundColor Green
        
        # Add custom columns
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "MessageId" -InternalName "MessageId" -Type Number -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ChannelUrl" -InternalName "ChannelUrl" -Type URL -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "DistributionStatus" -InternalName "DistributionStatus" -Type Choice -Choices @("Success","Failed","Pending") -DefaultValue "Pending" -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "DistributionTimestamp" -InternalName "DistributionTimestamp" -Type DateTime -AddToDefaultView
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ErrorMessage" -InternalName "ErrorMessage" -Type Note
        Add-PnPField -List "TeamsDistributionLogs" -DisplayName "ResponseData" -InternalName "ResponseData" -Type Note
        
        Write-Host "✅ Added custom columns to 'TeamsDistributionLogs' list" -ForegroundColor Green
    }
} catch {
    Write-Error "❌ Failed to create 'TeamsDistributionLogs' list: $($_.Exception.Message)"
}

# Create sample data if requested
if ($CreateSampleData) {
    Write-Host "`n🎯 Creating sample data..." -ForegroundColor Cyan
    
    try {
        # Get current date and future dates
        $today = Get-Date
        $tomorrow = $today.AddDays(1)
        $nextWeek = $today.AddDays(7)
        
        # Sample messages
        $sampleMessages = @(
            @{
                Title = "🚀 New Dashboard Features Released"
                MessageContent = "We're excited to announce the release of our new dashboard features! Check out the enhanced data visualization tools and improved performance."
                Priority = "High"
                ExpiryDate = $nextWeek
                TargetAudience = "All Users"
            },
            @{
                Title = "📅 Scheduled Maintenance Window"
                MessageContent = "Scheduled maintenance will occur this weekend from 2 AM to 6 AM. The system will be temporarily unavailable during this time."
                Priority = "Medium"
                ExpiryDate = $tomorrow
                TargetAudience = "All Users"
            },
            @{
                Title = "📊 Dashboard Tutorial Available"
                MessageContent = "New to the dashboard? Check out our comprehensive tutorial to learn about all the features and how to make the most of your analytics."
                Priority = "Low"
                ExpiryDate = $nextWeek
                TargetAudience = "New Users"
            }
        )
        
        foreach ($message in $sampleMessages) {
            Add-PnPListItem -List "AdaptiveCardMessages" -Values $message
            Write-Host "✅ Added sample message: $($message.Title)" -ForegroundColor Green
        }
        
        Write-Host "✅ Created sample data successfully" -ForegroundColor Green
    } catch {
        Write-Error "❌ Failed to create sample data: $($_.Exception.Message)"
    }
}

Write-Host "`n🎉 Setup completed successfully!" -ForegroundColor Green
Write-Host "Your Dashboard web part should now work properly." -ForegroundColor Green
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Go to your SharePoint site" -ForegroundColor White
Write-Host "2. Add the Dashboard web part to a page" -ForegroundColor White
Write-Host "3. The web part will now display messages from the 'AdaptiveCardMessages' list" -ForegroundColor White

Disconnect-PnPOnline
