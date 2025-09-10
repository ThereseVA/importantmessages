# PowerShell script to remove old SPFx web parts
# Run this script with SharePoint admin privileges

# Connect to your SharePoint site
$SiteUrl = "https://gustafkliniken.sharepoint.com/sites/appcatalog"

# Connect to SharePoint Online (you'll be prompted for credentials)
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "🔍 Searching for old SPFx apps..." -ForegroundColor Yellow

# Get all apps from App Catalog
$Apps = Get-PnPApp

# Display all apps so you can identify the old one
Write-Host "📋 Current apps in App Catalog:" -ForegroundColor Cyan
foreach ($App in $Apps) {
    Write-Host "  - $($App.Title) (ID: $($App.Id))" -ForegroundColor White
}

# Look for the old app (you'll need to identify it from the list above)
Write-Host ""
Write-Host "🎯 To remove the old app:" -ForegroundColor Green
Write-Host "1. Find the old 'spfx-adaptivecard-solution' app in the list above" -ForegroundColor White
Write-Host "2. Copy its ID" -ForegroundColor White
Write-Host "3. Run: Remove-PnPApp -Identity 'APP-ID-HERE'" -ForegroundColor White
Write-Host ""
Write-Host "Example:" -ForegroundColor Yellow
Write-Host "Remove-PnPApp -Identity 'c3d4e5f6-a7b8-9012-3456-789012cdefab'" -ForegroundColor Yellow

# Disconnect
Disconnect-PnPOnline

Write-Host "✅ Script completed. Manual removal required as shown above." -ForegroundColor Green
