# 📧 Simple Teams Channels List Creator
# Creates SharePoint list using different authentication methods

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$UseWebLogin,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateSampleData
)

Write-Host "📧 Creating TeamsChannels SharePoint List..." -ForegroundColor Green
Write-Host "🔗 Site: $SiteUrl" -ForegroundColor Cyan

try {
    # Try different connection methods
    if ($UseWebLogin) {
        Write-Host "🌐 Using web browser login..." -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
    } else {
        Write-Host "🔐 Using interactive login..." -ForegroundColor Cyan
        try {
            Connect-PnPOnline -Url $SiteUrl -Interactive
        } catch {
            Write-Host "⚠️ Interactive login failed, trying web login..." -ForegroundColor Yellow
            Connect-PnPOnline -Url $SiteUrl -UseWebLogin
        }
    }

    Write-Host "✅ Connected successfully!" -ForegroundColor Green

    # Check if list exists
    Write-Host "🔍 Checking if TeamsChannels list exists..." -ForegroundColor Cyan
    $existingList = Get-PnPList -Identity "TeamsChannels" -ErrorAction SilentlyContinue
    
    if ($existingList) {
        Write-Host "⚠️ TeamsChannels list already exists!" -ForegroundColor Yellow
        Write-Host "📋 List URL: $($existingList.DefaultViewUrl)" -ForegroundColor Gray
        $continue = Read-Host "Do you want to continue and add missing fields? (y/n)"
        if ($continue -ne 'y' -and $continue -ne 'Y') {
            Write-Host "❌ Operation cancelled" -ForegroundColor Red
            return
        }
    } else {
        # Create the list
        Write-Host "📝 Creating TeamsChannels list..." -ForegroundColor Cyan
        $newList = New-PnPList -Title "TeamsChannels" -Template GenericList -Url "Lists/TeamsChannels"
        Write-Host "✅ List created successfully!" -ForegroundColor Green
    }

    # Add fields one by one with error handling
    Write-Host "🔧 Adding list fields..." -ForegroundColor Cyan

    $fields = @(
        @{ Name = "ChannelName"; Type = "Text"; Required = $true; Description = "Name of the Teams channel" },
        @{ Name = "ChannelEmail"; Type = "Text"; Required = $true; Description = "Email address of the Teams channel" },
        @{ Name = "TeamName"; Type = "Text"; Required = $true; Description = "Name of the Teams team" },
        @{ Name = "Description"; Type = "Note"; Required = $false; Description = "Description of channel purpose" },
        @{ Name = "IsActive"; Type = "Boolean"; Required = $true; Description = "Whether to send messages to this channel" },
        @{ Name = "Department"; Type = "Text"; Required = $false; Description = "Department filter" },
        @{ Name = "MessageTypes"; Type = "Text"; Required = $false; Description = "Priority levels (High,Medium,Low)" }
    )

    foreach ($field in $fields) {
        try {
            Write-Host "  Adding field: $($field.Name)..." -ForegroundColor Gray
            
            $fieldXml = switch ($field.Type) {
                "Text" { 
                    if ($field.Required) {
                        "<Field Type='Text' DisplayName='$($field.Name)' Name='$($field.Name)' Required='TRUE' />"
                    } else {
                        "<Field Type='Text' DisplayName='$($field.Name)' Name='$($field.Name)' />"
                    }
                }
                "Note" { "<Field Type='Note' DisplayName='$($field.Name)' Name='$($field.Name)' />" }
                "Boolean" { "<Field Type='Boolean' DisplayName='$($field.Name)' Name='$($field.Name)' Required='TRUE'><Default>1</Default></Field>" }
            }
            
            Add-PnPFieldFromXml -List "TeamsChannels" -FieldXml $fieldXml -ErrorAction SilentlyContinue
            Write-Host "    ✅ $($field.Name) added" -ForegroundColor Green
        } catch {
            Write-Host "    ⚠️ $($field.Name) might already exist" -ForegroundColor Yellow
        }
    }

    # Update list description
    try {
        Set-PnPList -Identity "TeamsChannels" -Description "Configuration list for Teams channel email addresses. Used by Adaptive Cards solution to send messages to Teams channels via email."
        Write-Host "📝 List description updated" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Could not update description" -ForegroundColor Yellow
    }

    # Add sample data if requested
    if ($CreateSampleData) {
        Write-Host "`n📊 Adding sample data..." -ForegroundColor Cyan
        
        $sampleData = @(
            @{
                Title = "IT General Channel"
                ChannelName = "General"
                ChannelEmail = "REPLACE_WITH_ACTUAL_EMAIL@gustafkliniken.onmicrosoft.com"
                TeamName = "IT Department"
                Description = "General IT announcements and system updates"
                Department = "IT"
                MessageTypes = "High,Medium,Low"
                IsActive = $true
            },
            @{
                Title = "HR Announcements"
                ChannelName = "Announcements"
                ChannelEmail = "REPLACE_WITH_ACTUAL_EMAIL@gustafkliniken.onmicrosoft.com"
                TeamName = "HR Department"
                Description = "HR announcements and company news"
                Department = "HR"
                MessageTypes = "High,Medium"
                IsActive = $true
            }
        )

        foreach ($item in $sampleData) {
            try {
                Add-PnPListItem -List "TeamsChannels" -Values $item
                Write-Host "  ✅ Added: $($item.Title)" -ForegroundColor Green
            } catch {
                Write-Host "  ⚠️ Could not add: $($item.Title)" -ForegroundColor Yellow
            }
        }
    }

    # Get list URL for user
    $list = Get-PnPList -Identity "TeamsChannels"
    $listUrl = "$SiteUrl/Lists/TeamsChannels"

    Write-Host "`n🎉 SUCCESS! TeamsChannels list created!" -ForegroundColor Green
    Write-Host "📋 List URL: $listUrl" -ForegroundColor Cyan
    Write-Host "`n📧 NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. Open the list: $listUrl" -ForegroundColor White
    Write-Host "2. Get Teams channel emails (see guide below)" -ForegroundColor White
    Write-Host "3. Add real email addresses to replace sample data" -ForegroundColor White
    Write-Host "4. Test the email integration" -ForegroundColor White

    Write-Host "`n🔗 HOW TO GET TEAMS CHANNEL EMAILS:" -ForegroundColor Yellow
    Write-Host "For each Teams channel:" -ForegroundColor White
    Write-Host "  1. Open Microsoft Teams" -ForegroundColor Gray
    Write-Host "  2. Go to the channel (e.g., General, Announcements)" -ForegroundColor Gray  
    Write-Host "  3. Click '...' (three dots) next to channel name" -ForegroundColor Gray
    Write-Host "  4. Select 'Get email address'" -ForegroundColor Gray
    Write-Host "  5. Copy the email address" -ForegroundColor Gray
    Write-Host "  6. Paste it in the ChannelEmail column in SharePoint" -ForegroundColor Gray

} catch {
    Write-Host "❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`n💡 TROUBLESHOOTING:" -ForegroundColor Yellow
    Write-Host "Try running with -UseWebLogin parameter:" -ForegroundColor White
    Write-Host ".\create-teams-list.ps1 -SiteUrl '$SiteUrl' -UseWebLogin" -ForegroundColor Gray
}

Write-Host "`n✨ Script completed!" -ForegroundColor Green
