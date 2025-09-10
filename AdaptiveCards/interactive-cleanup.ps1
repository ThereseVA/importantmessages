# Quick script to identify and remove the old SPFx app
# This helps you find the exact app to remove

# Connect to your App Catalog
$SiteUrl = "https://gustafkliniken.sharepoint.com/sites/appcatalog"
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "🔍 Looking for SPFx Adaptive Card apps..." -ForegroundColor Yellow

# Get all apps and filter for our solution
$Apps = Get-PnPApp | Where-Object { 
    $_.Title -like "*adaptive*" -or 
    $_.Title -like "*dashboard*" -or
    $_.Title -like "*spfx*"
}

Write-Host ""
Write-Host "📋 Found these related apps:" -ForegroundColor Cyan

foreach ($App in $Apps) {
    $AppDetails = Get-PnPApp -Identity $App.Id
    Write-Host "  🎯 App: $($App.Title)" -ForegroundColor White
    Write-Host "     ID: $($App.Id)" -ForegroundColor Gray
    Write-Host "     Version: $($AppDetails.AppCatalogVersion)" -ForegroundColor Gray
    Write-Host "     Deployed: $($AppDetails.Deployed)" -ForegroundColor Gray
    Write-Host "     Installed: $($AppDetails.Installed)" -ForegroundColor Gray
    Write-Host ""
    
    # Ask if user wants to remove this app
    $Response = Read-Host "❓ Do you want to remove this app? (y/N)"
    if ($Response -eq "y" -or $Response -eq "Y") {
        try {
            Write-Host "🗑️ Removing app: $($App.Title)..." -ForegroundColor Red
            Remove-PnPApp -Identity $App.Id -Force
            Write-Host "✅ Successfully removed!" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Error removing app: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Write-Host "─────────────────────────────────────" -ForegroundColor Gray
}

Disconnect-PnPOnline
Write-Host "✅ Cleanup completed!" -ForegroundColor Green
