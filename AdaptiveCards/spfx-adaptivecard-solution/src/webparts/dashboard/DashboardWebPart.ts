import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './DashboardWebPart.module.scss';
import * as strings from 'DashboardWebPartStrings';
import { DashboardComponent } from './components/DashboardComponent';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IDashboardProps } from './components/IDashboardProps';
import { enhancedDataService } from '../../services/EnhancedDataService';

export interface IDashboardWebPartProps {
  title: string;
  description: string;
  dataSourceUrl: string;
  refreshInterval: number;
  showRefreshButton: boolean;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

  protected async onInit(): Promise<void> {
    console.log('🚀 DashboardWebPart.onInit() - Initializing Enhanced Data Service');
    
    // Set default values if not already configured
    if (!this.properties.dataSourceUrl) {
      this.properties.dataSourceUrl = 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken';
    }
    
    try {
      // Initialize the enhanced data service with Graph integration
      await enhancedDataService.initialize(this.context, this.properties.dataSourceUrl);
      console.log('✅ Enhanced Data Service initialized successfully for Dashboard');
      
      // Get current user info and check permissions
      const currentUser = enhancedDataService.getCurrentUser();
      if (currentUser) {
        console.log('👤 Dashboard user initialized:', {
          displayName: currentUser.spfx?.displayName || currentUser.graph?.displayName,
          groups: currentUser.groups?.length || 0,
          isManager: enhancedDataService.hasUserRole('manager'),
          isAdmin: enhancedDataService.hasUserRole('admin')
        });
      }
    } catch (error) {
      console.error('❌ Error initializing Enhanced Data Service for Dashboard:', error);
    }
    
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IDashboardProps> = React.createElement(
      DashboardComponent,
      {
        title: this.properties.title,
        description: this.properties.description,
        dataSourceUrl: this.properties.dataSourceUrl,
        refreshInterval: this.properties.refreshInterval,
        showRefreshButton: this.properties.showRefreshButton,
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('dataSourceUrl', {
                  label: strings.DataSourceUrlFieldLabel,
                  description: strings.DataSourceUrlFieldDescription
                })
              ]
            },
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyPaneTextField('refreshInterval', {
                  label: strings.RefreshIntervalFieldLabel,
                  description: strings.RefreshIntervalFieldDescription
                }),
                PropertyPaneToggle('showRefreshButton', {
                  label: strings.ShowRefreshButtonFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
