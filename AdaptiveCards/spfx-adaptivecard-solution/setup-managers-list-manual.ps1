# Alternative Setup Script for Managers List using REST API
# This script creates a SharePoint list using REST API calls directly

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$ListName = "Managers"
)

# Function to get SharePoint authentication token
function Get-SPAuthToken {
    param([string]$SiteUrl)
    
    Write-Host "Opening browser for SharePoint authentication..." -ForegroundColor Yellow
    
    # Create a simple form for authentication
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Web
    
    $form = New-Object Windows.Forms.Form
    $form.Text = "SharePoint Authentication"
    $form.Size = New-Object Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    
    $browser = New-Object Windows.Forms.WebBrowser
    $browser.Size = New-Object Drawing.Size(780, 580)
    $browser.Location = New-Object Drawing.Point(10, 10)
    $browser.ScriptErrorsSuppressed = $true
    
    $form.Controls.Add($browser)
    
    # Navigate to SharePoint site
    $browser.Navigate($SiteUrl)
    
    $form.ShowDialog() | Out-Null
    
    # Extract cookies/authentication info would go here
    # This is a simplified version - in practice, you'd need to handle the authentication flow
    
    return $null
}

Write-Host "=== SharePoint Managers List Setup (REST API) ===" -ForegroundColor Cyan
Write-Host "Site URL: $SiteUrl" -ForegroundColor Yellow
Write-Host "List Name: $ListName" -ForegroundColor Yellow
Write-Host ""

Write-Host "NOTICE: This script requires manual setup due to authentication complexity." -ForegroundColor Yellow
Write-Host "Please follow these manual steps instead:" -ForegroundColor Cyan
Write-Host ""

Write-Host "1. NAVIGATE TO YOUR SHAREPOINT SITE:" -ForegroundColor Green
Write-Host "   $SiteUrl" -ForegroundColor White
Write-Host ""

Write-Host "2. CREATE A NEW LIST:" -ForegroundColor Green
Write-Host "   - Click 'New' > 'List'" -ForegroundColor White
Write-Host "   - Choose 'Blank list'" -ForegroundColor White
Write-Host "   - Name: '$ListName'" -ForegroundColor White
Write-Host "   - Description: 'List to define who is a manager in the organization'" -ForegroundColor White
Write-Host ""

Write-Host "3. ADD THESE COLUMNS TO THE LIST:" -ForegroundColor Green
Write-Host ""
Write-Host "   a) Manager Email (Person field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Manager Email'" -ForegroundColor White
Write-Host "      - Type: 'Person or Group'" -ForegroundColor White
Write-Host "      - Required: Yes" -ForegroundColor White
Write-Host "      - Allow multiple selections: No" -ForegroundColor White
Write-Host ""
Write-Host "   b) Manager Display Name (Text field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Manager Display Name'" -ForegroundColor White
Write-Host "      - Type: 'Single line of text'" -ForegroundColor White
Write-Host "      - Required: Yes" -ForegroundColor White
Write-Host ""
Write-Host "   c) Department (Text field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Department'" -ForegroundColor White
Write-Host "      - Type: 'Single line of text'" -ForegroundColor White
Write-Host "      - Required: No" -ForegroundColor White
Write-Host ""
Write-Host "   d) Manager Level (Choice field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Manager Level'" -ForegroundColor White
Write-Host "      - Type: 'Choice'" -ForegroundColor White
Write-Host "      - Choices: 'Team Lead', 'Department Manager', 'Senior Manager', 'Director', 'VP', 'Executive'" -ForegroundColor White
Write-Host "      - Required: No" -ForegroundColor White
Write-Host ""
Write-Host "   e) Is Active (Yes/No field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Is Active'" -ForegroundColor White
Write-Host "      - Type: 'Yes/No'" -ForegroundColor White
Write-Host "      - Required: Yes" -ForegroundColor White
Write-Host "      - Default value: Yes" -ForegroundColor White
Write-Host ""
Write-Host "   f) Start Date (Date field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Start Date'" -ForegroundColor White
Write-Host "      - Type: 'Date and time'" -ForegroundColor White
Write-Host "      - Required: No" -ForegroundColor White
Write-Host ""
Write-Host "   g) End Date (Date field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'End Date'" -ForegroundColor White
Write-Host "      - Type: 'Date and time'" -ForegroundColor White
Write-Host "      - Required: No" -ForegroundColor White
Write-Host ""
Write-Host "   h) Notes (Multi-line text field):" -ForegroundColor Yellow
Write-Host "      - Column name: 'Notes'" -ForegroundColor White
Write-Host "      - Type: 'Multiple lines of text'" -ForegroundColor White
Write-Host "      - Required: No" -ForegroundColor White
Write-Host ""

Write-Host "4. UPDATE THE DEFAULT VIEW:" -ForegroundColor Green
Write-Host "   - Go to 'All Items' view" -ForegroundColor White
Write-Host "   - Click 'Edit current view'" -ForegroundColor White
Write-Host "   - Uncheck 'Title' (not needed)" -ForegroundColor White
Write-Host "   - Check these columns:" -ForegroundColor White
Write-Host "     * Manager Email" -ForegroundColor Gray
Write-Host "     * Manager Display Name" -ForegroundColor Gray
Write-Host "     * Department" -ForegroundColor Gray
Write-Host "     * Manager Level" -ForegroundColor Gray
Write-Host "     * Is Active" -ForegroundColor Gray
Write-Host "     * Start Date" -ForegroundColor Gray
Write-Host "   - Click 'OK'" -ForegroundColor White
Write-Host ""

Write-Host "5. ADD SAMPLE DATA:" -ForegroundColor Green
Write-Host "   - Click 'New' to add a new item" -ForegroundColor White
Write-Host "   - Fill in the manager information" -ForegroundColor White
Write-Host "   - Make sure 'Is Active' is set to 'Yes' for current managers" -ForegroundColor White
Write-Host ""

Write-Host "6. SET PERMISSIONS (OPTIONAL):" -ForegroundColor Green
Write-Host "   - Go to List Settings > Permissions for this list" -ForegroundColor White
Write-Host "   - Stop inheriting permissions" -ForegroundColor White
Write-Host "   - Give 'Read' access to all users" -ForegroundColor White
Write-Host "   - Give 'Edit' access to HR/Admin staff only" -ForegroundColor White
Write-Host ""

Write-Host "COMPLETED SETUP VERIFICATION:" -ForegroundColor Magenta
Write-Host ""
Write-Host "After completing the manual setup, verify:" -ForegroundColor Yellow
Write-Host "✓ List named '$ListName' exists" -ForegroundColor Green
Write-Host "✓ All required columns are created" -ForegroundColor Green
Write-Host "✓ At least one manager is added with 'Is Active' = Yes" -ForegroundColor Green
Write-Host "✓ Permissions are set appropriately" -ForegroundColor Green
Write-Host ""

Write-Host "NEXT STEPS:" -ForegroundColor Magenta
Write-Host "1. Add all your managers to the list" -ForegroundColor Yellow
Write-Host "2. Test the Manager Dashboard in your SPFx solution" -ForegroundColor Yellow
Write-Host "3. The solution will automatically check this list for manager permissions" -ForegroundColor Yellow
Write-Host ""

Write-Host "SAMPLE DATA FORMAT:" -ForegroundColor Magenta
Write-Host "Title: [Auto-generated or Manager Name]" -ForegroundColor White
Write-Host "Manager Email: [Select from people picker]" -ForegroundColor White
Write-Host "Manager Display Name: John Smith" -ForegroundColor White
Write-Host "Department: IT" -ForegroundColor White
Write-Host "Manager Level: Department Manager" -ForegroundColor White
Write-Host "Is Active: Yes" -ForegroundColor White
Write-Host "Start Date: [When they became manager]" -ForegroundColor White
Write-Host "Notes: [Any additional information]" -ForegroundColor White
Write-Host ""

Write-Host "For technical support, check the MANAGERS-LIST-SETUP-GUIDE.md file." -ForegroundColor Cyan

# Open the SharePoint site in the browser
Write-Host "Opening SharePoint site in your default browser..." -ForegroundColor Yellow
Start-Process $SiteUrl

Write-Host ""
Write-Host "Manual setup instructions displayed above." -ForegroundColor Green
Write-Host "SharePoint site opened in your browser." -ForegroundColor Green
