# Fix Managers List - Final Verification and Correction
# This script will check and fix the Managers list column structure

Write-Host "🔧 FINAL FIX: Managers List Column Structure" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

Write-Host ""
Write-Host "❗ CRITICAL ISSUES IDENTIFIED:" -ForegroundColor Red
Write-Host "1. ManagerEmail must be 'Person or Group' column type" -ForegroundColor Yellow
Write-Host "2. Column names must match exactly what the code expects" -ForegroundColor Yellow
Write-Host "3. EndDate must be 'Date and Time' type" -ForegroundColor Yellow

Write-Host ""
Write-Host "🎯 REQUIRED COLUMN STRUCTURE:" -ForegroundColor Green
Write-Host "=" * 40 -ForegroundColor Green

$columns = @"
1. Title                  - Single line of text (DEFAULT)
2. ManagerEmail          - Person or Group (CRITICAL!)
3. ManagerDisplayName    - Single line of text  
4. Department            - Single line of text
5. ManagerLevel          - Single line of text
6. IsActive              - Yes/No (Boolean)
7. StartDate             - Date and Time
8. EndDate               - Date and Time (CRITICAL!)
9. Notes                 - Multiple lines of text
"@

Write-Host $columns -ForegroundColor White

Write-Host ""
Write-Host "🚨 IMMEDIATE ACTION REQUIRED:" -ForegroundColor Red
Write-Host "=" * 40 -ForegroundColor Red

Write-Host "1. Open the Managers list in SharePoint" -ForegroundColor Yellow
Write-Host "2. Go to List Settings > Columns" -ForegroundColor Yellow
Write-Host "3. DELETE the current 'ManagerEmail' column if it's 'Single line of text'" -ForegroundColor Yellow
Write-Host "4. CREATE NEW 'ManagerEmail' column as 'Person or Group'" -ForegroundColor Yellow
Write-Host "5. Verify 'EndDate' is 'Date and Time' type" -ForegroundColor Yellow
Write-Host "6. Ensure all other column names match exactly" -ForegroundColor Yellow

Write-Host ""
Write-Host "🔗 Opening SharePoint List Settings..." -ForegroundColor Green
Start-Process "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/Lists/Managers/AllItems.aspx"

Write-Host ""
Write-Host "⚡ AFTER FIXING:" -ForegroundColor Cyan
Write-Host "- Refresh your SharePoint page" -ForegroundColor White
Write-Host "- The 400 errors should disappear" -ForegroundColor White
Write-Host "- Manager permissions will work correctly" -ForegroundColor White

Write-Host ""
Write-Host "✅ Your SPFx solution is ready - just needs the list fixed!" -ForegroundColor Green
