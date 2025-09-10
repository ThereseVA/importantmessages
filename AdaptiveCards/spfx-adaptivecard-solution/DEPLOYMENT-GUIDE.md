# 🚀 SPFx Adaptive Cards Solution - Deployment Guide

## SharePoint Online Deployment Steps

### 1. Package for Production
```powershell
# Build production bundle
gulp build --ship

# Package solution
gulp bundle --ship
gulp package-solution --ship
```

### 2. Upload to App Catalog
1. Navigate to your SharePoint Admin Center
2. Go to "More features" > "Apps" > "App Catalog"
3. Upload the `.sppkg` file from `sharepoint/solution/`
4. Click "Deploy" and trust the solution

### 3. Add to SharePoint Site
1. Go to your target SharePoint site
2. Settings gear > "Add an app"
3. Find "spfx-adaptivecard-solution" and add it
4. Add the web part to any page

### 4. Configure Card Data Source
- Update the Card JSON URL to point to your production data source
- Ensure CORS is properly configured for external APIs
- Test with real SharePoint data

## Alternative Deployment Options

### Microsoft Teams Deployment
- The solution can be deployed to Teams as a Teams app
- Use the Teams manifest in the `teams/` folder
- Package and upload to Teams App Catalog

### Local Development Testing
- Use `gulp serve` for local development server
- Test with SharePoint Workbench
- Use the testing environment for component validation

## Configuration Options

### Environment-Specific Settings
- Development: `http://localhost:3000/api/test-card`
- Staging: `https://your-staging-api.com/card-data`
- Production: `https://your-production-api.com/card-data`

### Security Considerations
- Ensure API endpoints support SharePoint domain CORS
- Implement proper authentication for data sources
- Follow SharePoint security best practices
