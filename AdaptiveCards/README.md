# SharePoint Framework Adaptive Cards Solution

A comprehensive SharePoint Framework solution that demonstrates Adaptive Cards integration with dashboard functionality.

## Features

- **Adaptive Card Viewer**: Renders dynamic Adaptive Cards with rich content
- **Dashboard Interface**: Provides a centralized view for managing and displaying data
- **Data Services**: Handles backend communication with SharePoint lists and external APIs
- **Responsive Design**: Mobile-friendly UI components

## Project Structure

```
spfx-adaptivecard-solution/
├── src/
│   ├── webparts/
│   │   ├── adaptiveCardViewer/     # Adaptive Card rendering component
│   │   └── dashboard/              # Dashboard web part
│   └── services/
│       └── DataService.ts          # Backend communication service
├── config/                         # SPFx configuration files
└── package.json                    # Project dependencies
```

## Getting Started

### Prerequisites

- Node.js (v18.17.1 or higher)
- SharePoint Framework development environment
- Yeoman and Gulp CLI

### Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```

### Development

1. Start the local development server:
   ```bash
   npm run serve
   ```

2. Test your web part in the SharePoint Workbench

### Building for Production

1. Bundle and package the solution:
   ```bash
   npm run package-solution
   ```

2. Deploy the `.sppkg` file to your SharePoint App Catalog

## Web Parts

### Adaptive Card Viewer
- Renders Adaptive Cards from various data sources
- Supports interactive elements and actions
- Customizable styling and theming

### Dashboard
- Centralized data visualization
- Real-time updates from SharePoint lists
- Configurable layout and widgets

## Data Services

The `DataService` class provides:
- SharePoint list operations (CRUD)
- External API integration
- Caching and performance optimization
- Error handling and logging

## Configuration

Modify `config/config.json` to customize:
- API endpoints
- SharePoint list configurations
- Feature flags
- Environment settings

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

This project is licensed under the MIT License.
