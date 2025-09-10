# 📧 Teams Channels Configuration Setup
# Creates SharePoint list to store Teams channel email addresses

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateSampleData
)

Write-Host "📧 Setting up Teams Channels configuration list..." -ForegroundColor Green

try {
    # Connect to SharePoint
    Write-Host "🔗 Connecting to SharePoint..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive

    # Check if list already exists
    $existingList = Get-PnPList -Identity "TeamsChannels" -ErrorAction SilentlyContinue
    
    if ($existingList) {
        Write-Host "⚠️ TeamsChannels list already exists" -ForegroundColor Yellow
        $response = Read-Host "Do you want to add missing fields? (y/n)"
        if ($response -ne 'y') {
            Write-Host "❌ Setup cancelled" -ForegroundColor Red
            exit
        }
    } else {
        # Create the list
        Write-Host "📝 Creating TeamsChannels list..." -ForegroundColor Cyan
        New-PnPList -Title "TeamsChannels" -Template GenericList -Url "Lists/TeamsChannels"
    }

    # Add/Update fields
    Write-Host "🔧 Adding/updating list fields..." -ForegroundColor Cyan

    try {
        # ChannelName - Single line of text (required)
        Add-PnPField -List "TeamsChannels" -DisplayName "ChannelName" -InternalName "ChannelName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
        
        # ChannelEmail - Single line of text (required)
        Add-PnPField -List "TeamsChannels" -DisplayName "ChannelEmail" -InternalName "ChannelEmail" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
        
        # TeamName - Single line of text (required)
        Add-PnPField -List "TeamsChannels" -DisplayName "TeamName" -InternalName "TeamName" -Type Text -Required -AddToDefaultView -ErrorAction SilentlyContinue
        
        # Description - Multiple lines of text (optional)
        Add-PnPField -List "TeamsChannels" -DisplayName "Description" -InternalName "Description" -Type Note -AddToDefaultView -ErrorAction SilentlyContinue
        
        # IsActive - Boolean field (required, default true)
        Add-PnPField -List "TeamsChannels" -DisplayName "IsActive" -InternalName "IsActive" -Type Boolean -DefaultValue $true -Required -AddToDefaultView -ErrorAction SilentlyContinue
        
        # Department - Single line of text (optional)
        Add-PnPField -List "TeamsChannels" -DisplayName "Department" -InternalName "Department" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
        
        # MessageTypes - Single line of text (optional)
        Add-PnPField -List "TeamsChannels" -DisplayName "MessageTypes" -InternalName "MessageTypes" -Type Text -AddToDefaultView -ErrorAction SilentlyContinue
        
        Write-Host "✅ Fields created successfully" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Some fields might already exist: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Set list description
    try {
        Set-PnPList -Identity "TeamsChannels" -Description "Configuration list for Teams channel email addresses. Used by the Adaptive Cards solution to automatically send messages to configured Teams channels."
        Write-Host "📝 List description updated" -ForegroundColor Green
    } catch {
        Write-Host "⚠️ Could not update list description: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Add sample data if requested
    if ($CreateSampleData) {
        Write-Host "`n📊 Creating sample Teams channel data..." -ForegroundColor Cyan
        try {
            # Sample channels - Update these with your actual Teams channel emails
            $sampleChannels = @(
                @{
                    Title = "IT Department - General"
                    ChannelName = "General"
                    ChannelEmail = "general_abc123@gustafkliniken.onmicrosoft.com"
                    TeamName = "IT Department"
                    Description = "General IT announcements and updates"
                    Department = "IT"
                    MessageTypes = "High,Medium,Low"
                    IsActive = $true
                },
                @{
                    Title = "HR Department - Announcements"
                    ChannelName = "Announcements"
                    ChannelEmail = "announcements_def456@gustafkliniken.onmicrosoft.com"
                    TeamName = "HR Department"
                    Description = "HR announcements and company news"
                    Department = "HR"
                    MessageTypes = "High,Medium"
                    IsActive = $true
                },
                @{
                    Title = "Management - Urgent Only"
                    ChannelName = "Urgent Communications"
                    ChannelEmail = "urgent_ghi789@gustafkliniken.onmicrosoft.com"
                    TeamName = "Management Team"
                    Description = "Urgent communications for management team"
                    Department = "Management"
                    MessageTypes = "High"
                    IsActive = $true
                },
                @{
                    Title = "All Staff - General"
                    ChannelName = "Company News"
                    ChannelEmail = "news_jkl012@gustafkliniken.onmicrosoft.com"
                    TeamName = "All Staff"
                    Description = "General company news and updates for all staff"
                    Department = "All"
                    MessageTypes = "High,Medium,Low"
                    IsActive = $true
                }
            )
            
            foreach ($channel in $sampleChannels) {
                try {
                    # Check if channel already exists
                    $existingChannel = Get-PnPListItem -List "TeamsChannels" -Query "<View><Query><Where><Eq><FieldRef Name='ChannelEmail'/><Value Type='Text'>$($channel.ChannelEmail)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
                    
                    if ($existingChannel) {
                        Write-Host "⚠️ Channel already exists: $($channel.ChannelEmail)" -ForegroundColor Yellow
                    } else {
                        Add-PnPListItem -List "TeamsChannels" -Values $channel
                        Write-Host "✅ Added channel: $($channel.TeamName) - $($channel.ChannelName)" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "⚠️ Could not add channel $($channel.ChannelEmail): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "⚠️ Error creating sample data: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Write-Host "`n✅ TeamsChannels list setup completed!" -ForegroundColor Green
    
    # Display summary
    Write-Host "`n📋 LIST CONFIGURATION:" -ForegroundColor Cyan
    Write-Host "   • List Name: TeamsChannels" -ForegroundColor Gray
    Write-Host "   • ChannelName (Text, Required) - Name of the Teams channel" -ForegroundColor Gray
    Write-Host "   • ChannelEmail (Text, Required) - Email address of the Teams channel" -ForegroundColor Gray
    Write-Host "   • TeamName (Text, Required) - Name of the Teams team" -ForegroundColor Gray
    Write-Host "   • Description (Note, Optional) - Description of the channel purpose" -ForegroundColor Gray
    Write-Host "   • IsActive (Boolean, Required) - Whether to send messages to this channel" -ForegroundColor Gray
    Write-Host "   • Department (Text, Optional) - Department filter for targeted messages" -ForegroundColor Gray
    Write-Host "   • MessageTypes (Text, Optional) - Comma-separated priority levels (High,Medium,Low)" -ForegroundColor Gray

    Write-Host "`n🔗 HOW TO GET CHANNEL EMAILS:" -ForegroundColor Cyan
    Write-Host "1. Go to Teams channel" -ForegroundColor White
    Write-Host "2. Click '...' (more options)" -ForegroundColor White
    Write-Host "3. Choose 'Get email address'" -ForegroundColor White
    Write-Host "4. Copy the email address" -ForegroundColor White
    Write-Host "5. Add it to the TeamsChannels list" -ForegroundColor White

    Write-Host "`n🚀 NEXT STEPS:" -ForegroundColor Cyan
    Write-Host "1. 📧 Add your real Teams channel emails to the 'TeamsChannels' list" -ForegroundColor White
    Write-Host "2. 🔄 Deploy the updated SPFx solution with TeamsChannelService" -ForegroundColor White
    Write-Host "3. ✅ Test email sending to Teams channels" -ForegroundColor White
    Write-Host "4. 🎯 The system will automatically send to configured channels based on priority/department" -ForegroundColor White

    Write-Host "`n💡 USAGE EXAMPLES:" -ForegroundColor Cyan
    Write-Host "• Send to all channels: TeamsChannelService.sendToConfiguredChannels(title, message)" -ForegroundColor White
    Write-Host "• Send to specific department: TeamsChannelService.sendToConfiguredChannels(title, message, priority, 'IT')" -ForegroundColor White
    Write-Host "• Send only high priority: Use MessageTypes filter in the list" -ForegroundColor White
    Write-Host "• Temporarily disable channel: Set IsActive to False" -ForegroundColor White

} catch {
    Write-Host "❌ Error setting up TeamsChannels list: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "💡 Make sure you have proper permissions to create lists in SharePoint" -ForegroundColor Yellow
}

Write-Host "`n🎉 Setup complete! Your Teams channels are now ready for email integration." -ForegroundColor Green
