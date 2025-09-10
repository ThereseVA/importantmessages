# Email Mismatch Diagnostic Tool
# This script helps identify why your manager access isn't working

param(
    [string]$SiteUrl = "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken"
)

Write-Host "🔍 EMAIL MISMATCH DIAGNOSTIC TOOL" -ForegroundColor Yellow
Write-Host "=================================" -ForegroundColor Yellow
Write-Host ""

Write-Host "Your Current Login Email: therese.almesjo@gustafkliniken.se" -ForegroundColor Green
Write-Host ""

Write-Host "📋 STEPS TO DIAGNOSE:" -ForegroundColor Cyan
Write-Host "1. Check the Managers list REST API response in your browser" -ForegroundColor White
Write-Host "2. Look for your entry in the JSON response" -ForegroundColor White  
Write-Host "3. Find the 'ManagerEmail' object and check the 'EMail' field" -ForegroundColor White
Write-Host "4. Compare it with your login email above" -ForegroundColor White
Write-Host ""

Write-Host "🔧 COMMON FIXES:" -ForegroundColor Magenta
Write-Host "1. DOMAIN MISMATCH: If the email shows .com instead of .se (or vice versa)" -ForegroundColor White
Write-Host "   - Edit your Managers list entry" -ForegroundColor White
Write-Host "   - Remove yourself from ManagerEmail field" -ForegroundColor White
Write-Host "   - Re-add using People Picker with correct domain" -ForegroundColor White
Write-Host ""
Write-Host "2. DIFFERENT EMAIL: If it shows a completely different email" -ForegroundColor White
Write-Host "   - Your profile might have multiple emails" -ForegroundColor White
Write-Host "   - Remove and re-add yourself in ManagerEmail field" -ForegroundColor White
Write-Host ""
Write-Host "3. MISSING ENTRY: If you don't see your entry at all" -ForegroundColor White
Write-Host "   - Add yourself to the Managers list" -ForegroundColor White
Write-Host "   - Set IsActive to Yes" -ForegroundColor White
Write-Host ""

Write-Host "🌐 Opening diagnostic URLs..." -ForegroundColor Yellow
Write-Host ""

# Open the Managers list with detailed email information
$managersUrl = "${SiteUrl}/_api/web/lists/getbytitle('Managers')/items?`$expand=ManagerEmail&`$select=Id,Title,ManagerEmail/EMail,ManagerEmail/Title,ManagerDisplayName,IsActive"
Write-Host "Opening Managers List (JSON): $managersUrl" -ForegroundColor Cyan
Start-Process $managersUrl

Start-Sleep -Seconds 2

# Open the current user API
$currentUserUrl = "${SiteUrl}/_api/web/currentuser"
Write-Host "Opening Current User Info (JSON): $currentUserUrl" -ForegroundColor Cyan
Start-Process $currentUserUrl

Write-Host ""
Write-Host "✅ WHAT TO DO NEXT:" -ForegroundColor Green
Write-Host "1. Compare the emails in both browser tabs" -ForegroundColor White
Write-Host "2. If they don't match exactly, fix the ManagerEmail field" -ForegroundColor White
Write-Host "3. Test access again after fixing" -ForegroundColor White
Write-Host ""
Write-Host "💡 TIP: Copy both email values and compare them character by character" -ForegroundColor Yellow
