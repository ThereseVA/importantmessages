# SharePoint REST API Diagnostic Script
# This will help identify why the 400 error is still occurring

Write-Host "🔧 SHAREPOINT REST API DIAGNOSTIC" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan

Write-Host ""
Write-Host "❗ 400 Error Still Occurring - Possible Causes:" -ForegroundColor Red
Write-Host "=" * 45 -ForegroundColor Red

Write-Host ""
Write-Host "1. LIST NAME ISSUE:" -ForegroundColor Yellow
Write-Host "   - List might be called 'Managers' but internal name different" -ForegroundColor White
Write-Host "   - Try: 'Managers', 'ManagersList', or check the URL" -ForegroundColor White

Write-Host ""
Write-Host "2. COLUMN INTERNAL NAMES:" -ForegroundColor Yellow
Write-Host "   - Display name 'ManagerEmail' might have internal name 'ManagerEmail0'" -ForegroundColor White
Write-Host "   - SharePoint sometimes adds numbers to duplicate names" -ForegroundColor White

Write-Host ""
Write-Host "3. MISSING DATA:" -ForegroundColor Yellow
Write-Host "   - List exists but has no items (empty)" -ForegroundColor White
Write-Host "   - Required columns might be missing values" -ForegroundColor White

Write-Host ""
Write-Host "4. PERMISSIONS:" -ForegroundColor Yellow
Write-Host "   - User might not have read access to the list" -ForegroundColor White
Write-Host "   - List permissions might be restricted" -ForegroundColor White

Write-Host ""
Write-Host "🔍 DIAGNOSTIC STEPS:" -ForegroundColor Green
Write-Host "=" * 25 -ForegroundColor Green

Write-Host ""
Write-Host "Step 1: Test basic list access" -ForegroundColor Cyan
Write-Host "URL: https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')" -ForegroundColor White

Write-Host ""
Write-Host "Step 2: Test simple items query" -ForegroundColor Cyan  
Write-Host "URL: https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')/items" -ForegroundColor White

Write-Host ""
Write-Host "Step 3: Test with minimal select" -ForegroundColor Cyan
Write-Host "URL: https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')/items?`$select=Id,Title" -ForegroundColor White

Write-Host ""
Write-Host "Step 4: Test ManagerEmail expansion" -ForegroundColor Cyan
Write-Host "URL: https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')/items?`$select=Id,ManagerEmail/EMail&`$expand=ManagerEmail" -ForegroundColor White

Write-Host ""
Write-Host "🚨 IMMEDIATE FIXES TO TRY:" -ForegroundColor Red
Write-Host "=" * 30 -ForegroundColor Red

Write-Host ""
Write-Host "1. ADD A TEST ITEM to the Managers list:" -ForegroundColor Yellow
Write-Host "   - Go to the list and click 'New'" -ForegroundColor White
Write-Host "   - Fill in at least: Title, ManagerEmail (select a person), IsActive=Yes" -ForegroundColor White

Write-Host ""
Write-Host "2. CHECK LIST PERMISSIONS:" -ForegroundColor Yellow
Write-Host "   - List Settings > Permissions" -ForegroundColor White
Write-Host "   - Ensure current user has Read access" -ForegroundColor White

Write-Host ""
Write-Host "3. VERIFY INTERNAL COLUMN NAMES:" -ForegroundColor Yellow
Write-Host "   - List Settings > [Column Name] > Check URL for internal name" -ForegroundColor White

Write-Host ""
Write-Host "🔗 Opening diagnostic URLs..." -ForegroundColor Green

# Test basic list access
Write-Host ""
Write-Host "Testing: Basic list access..." -ForegroundColor Cyan
Start-Process "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')"

Start-Sleep 2

# Test items access  
Write-Host "Testing: Items access..." -ForegroundColor Cyan
Start-Process "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')/items"

Start-Sleep 2

# Test minimal select
Write-Host "Testing: Minimal select..." -ForegroundColor Cyan
Start-Process "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/_api/web/lists/getbytitle('Managers')/items?`$select=Id,Title"

Write-Host ""
Write-Host "✅ CHECK EACH URL - Look for:" -ForegroundColor Green
Write-Host "   - 200 OK = Success" -ForegroundColor White
Write-Host "   - 400 Bad Request = Problem with that specific query" -ForegroundColor White
Write-Host "   - 404 Not Found = List does not exist or wrong name" -ForegroundColor White
Write-Host "   - 403 Forbidden = Permission issue" -ForegroundColor White
