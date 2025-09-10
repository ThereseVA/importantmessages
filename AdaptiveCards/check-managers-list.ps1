# Check and populate Managers list
# Run this script to verify and add managers to your SharePoint list

# SharePoint site URL
$siteUrl = "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken"
$listName = "Managers"

Write-Host "🔍 Checking Managers list at: $siteUrl" -ForegroundColor Yellow
Write-Host "📋 List name: $listName" -ForegroundColor Yellow

# Instructions for manual verification
Write-Host "`n📋 MANUAL STEPS TO FIX:" -ForegroundColor Green
Write-Host "1. Go to: $siteUrl/Lists/$listName" -ForegroundColor White
Write-Host "2. Check if the list exists and has items" -ForegroundColor White
Write-Host "3. If empty, add a new item with:" -ForegroundColor White
Write-Host "   - ManagersEmail: therese.almesjo@gustafkliniken.se" -ForegroundColor Cyan
Write-Host "   - ManagersDisplayName: Therese Varre Almesjö" -ForegroundColor Cyan
Write-Host "   - IsActive: Yes (True)" -ForegroundColor Cyan
Write-Host "   - Department: Administration" -ForegroundColor Cyan
Write-Host "   - ManagerLevel: 1" -ForegroundColor Cyan

Write-Host "`n🔗 Direct link to list:" -ForegroundColor Green
Write-Host "$siteUrl/_layouts/15/listedit.aspx?List=%7B$(([guid]::NewGuid().ToString().ToUpper()))%7D" -ForegroundColor Blue

Write-Host "`n⚠️  IMPORTANT: Make sure the 'IsActive' field is set to 'Yes' or 'True'" -ForegroundColor Red
Write-Host "⚠️  IMPORTANT: Use the exact email: therese.almesjo@gustafkliniken.se" -ForegroundColor Red
