import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './AdaptiveCardViewerWebPart.module.scss';
import * as strings from 'AdaptiveCardViewerWebPartStrings';
import { AdaptiveCardComponent } from './components/AdaptiveCardComponent';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IAdaptiveCardProps } from './components/IAdaptiveCardProps';
import { enhancedDataService } from '../../services/EnhancedDataService';

export interface IAdaptiveCardViewerWebPartProps {
  cardJsonUrl: string;
  title: string;
  enableActions: boolean;
  cardSource: string;
}

export default class AdaptiveCardViewerWebPart extends BaseClientSideWebPart<IAdaptiveCardViewerWebPartProps> {

  protected async onInit(): Promise<void> {
    console.log('🚀 AdaptiveCardViewerWebPart.onInit() - Initializing Enhanced Data Service');
    
    try {
      // Initialize the enhanced data service with Graph integration
      await enhancedDataService.initialize(this.context);
      console.log('✅ Enhanced Data Service initialized successfully');
      
      // Get current user info for logging
      const currentUser = enhancedDataService.getCurrentUser();
      if (currentUser) {
        console.log('👤 Current user initialized:', {
          displayName: currentUser.spfx?.displayName || currentUser.graph?.displayName,
          groups: currentUser.groups?.length || 0,
          hasPhoto: currentUser.hasPhoto
        });
      }
    } catch (error) {
      console.error('❌ Error initializing Enhanced Data Service:', error);
    }

    return super.onInit();
  }

  public render(): void {
    console.log('🚀 AdaptiveCardViewerWebPart.render() called');
    console.log('🔧 Properties:', this.properties);
    console.log('🌐 Context available:', !!this.context);
    console.log('📱 Display mode:', this.displayMode);
    console.log('🎯 DOM element:', this.domElement);
    
    // Determine card URL based on source selection
    let cardJsonUrl: string;
    
    // Check for URL parameter override first
    const urlParams = new URLSearchParams(window.location.search);
    const componentParam = urlParams.get('component');
    
    let cardSource = this.properties.cardSource || 'manager-dashboard';
    
    // Override cardSource with URL parameter if present
    if (componentParam) {
      switch (componentParam) {
        case 'teams-message-creator':
        case 'manager-dashboard':
        case 'message-list-diagnostic':
          cardSource = componentParam;
          console.log('📋 Card source overridden by URL parameter:', cardSource);
          break;
        default:
          console.log('📋 Unknown component parameter, using default:', cardSource);
      }
    } else {
      console.log('📋 Card source selected:', cardSource);
    }
    
    switch (cardSource) {
      case 'manager-dashboard':
        cardJsonUrl = 'component:manager-dashboard';
        break;
      case 'teams-message-creator':
        cardJsonUrl = 'component:teams-message-creator';
        break;
      case 'message-list-diagnostic':
        cardJsonUrl = 'component:message-list-diagnostic';
        break;
      case 'asset-sample':
        cardJsonUrl = 'asset:sample-card';
        break;
      case 'template-sample':
        cardJsonUrl = 'template:sample';
        break;
      case 'template-dashboard':
        cardJsonUrl = 'template:dashboard';
        break;
      case 'asset-project':
        cardJsonUrl = 'asset:project-status';
        break;
      case 'asset-team':
        cardJsonUrl = 'asset:team-notification';
        break;
      case 'asset-sales':
        cardJsonUrl = 'asset:sales-dashboard';
        break;
      case 'custom':
        cardJsonUrl = this.properties.cardJsonUrl || 'asset:sample-card';
        break;
      default:
        cardJsonUrl = 'component:manager-dashboard';
        break;
    }
    
    console.log('🔗 Final cardJsonUrl:', cardJsonUrl);
    
    try {
      const element: React.ReactElement<IAdaptiveCardProps> = React.createElement(
        AdaptiveCardComponent,
        {
          cardJsonUrl: cardJsonUrl,
          title: this.properties.title || 'Adaptive Card Demo',
          enableActions: this.properties.enableActions !== false,
          context: this.context,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          }
        }
      );
  
      console.log('⚛️ React element created successfully:', element);
      console.log('🎨 About to render to DOM element:', this.domElement);
      
      ReactDom.render(element, this.domElement);
      console.log('✅ ReactDom.render completed successfully');
      
    } catch (error) {
      console.error('❌ Error in render method:', error);
      // Render error message directly to DOM
      if (this.domElement) {
        this.domElement.innerHTML = `
          <div style="color: red; padding: 20px; border: 1px solid red; margin: 10px;">
            <h3>Web Part Render Error</h3>
            <p>Error: ${error.message || error}</p>
            <p>Please check browser console for details.</p>
          </div>
        `;
      }
    }
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
                PropertyPaneDropdown('cardSource', {
                  label: 'Select Tool',
                  options: [
                    { key: 'manager-dashboard', text: '🎛️ Manager Dashboard (CREATE MESSAGES)' },
                    { key: 'teams-message-creator', text: '� Teams Message Creator' },
                    { key: 'message-list-diagnostic', text: '🔍 Message List Diagnostic' },
                    { key: 'asset-sample', text: '📋 Sample Card Demo' }
                  ],
                  selectedKey: this.properties.cardSource || 'manager-dashboard'
                }),
                PropertyPaneTextField('cardJsonUrl', {
                  label: strings.CardJsonUrlFieldLabel,
                  description: strings.CardJsonUrlFieldDescription,
                  disabled: this.properties.cardSource !== 'custom'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
