import * as React from 'react';
import styles from './AdaptiveCardComponent.module.scss';
import { IAdaptiveCardProps } from './IAdaptiveCardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import * as AdaptiveCards from 'adaptivecards';
import { cardTemplates } from '../models/CardTemplates';
import { TeamsMessageCreator } from './TeamsMessageCreator';
import { ManagerDashboard } from './ManagerDashboard';

export interface IAdaptiveCardState {
  cardData: any;
  loading: boolean;
  error: string | null;
}

export class AdaptiveCardComponent extends React.Component<IAdaptiveCardProps, IAdaptiveCardState> {
  private cardContainer: React.RefObject<HTMLDivElement>;

  constructor(props: IAdaptiveCardProps) {
    super(props);
    console.log('🏗️ AdaptiveCardComponent constructor called with props:', props);
    this.cardContainer = React.createRef();
    this.state = {
      cardData: null,
      loading: false,
      error: null
    };
    console.log('🏗️ AdaptiveCardComponent constructor completed');
  }

  public async componentDidMount(): Promise<void> {
    console.log('🚀 AdaptiveCardComponent v1.0.34.0 - Component mounted with Teams multi-site support');
    console.log('📊 Component state:', this.state);
    console.log('📋 Component props:', this.props);
    console.log('🔗 Card JSON URL:', this.props.cardJsonUrl);
    
    // Initialize enhanced data service
    try {
      console.log('🔧 Initializing Enhanced DataService...');
      await enhancedDataService.initialize(this.props.context, this.props.context?.pageContext?.web?.absoluteUrl);
      console.log('✅ Enhanced DataService initialized successfully');
    } catch (error) {
      console.error('❌ Error initializing Enhanced DataService:', error);
    }
    
    if (!this.props.cardJsonUrl) {
      console.log('⚠️ No cardJsonUrl provided, rendering default card');
      // Render default card if no URL is configured
      setTimeout(() => this.renderAdaptiveCard(this.getDefaultCard()), 100);
    } else {
      console.log('🎯 Loading card from URL:', this.props.cardJsonUrl);
      // Load card from URL
      this.loadCard();
    }
  }

  public componentDidUpdate(prevProps: IAdaptiveCardProps): void {
    if (prevProps.cardJsonUrl !== this.props.cardJsonUrl && this.props.cardJsonUrl) {
      this.loadCard();
    }
  }

  private async loadCard(): Promise<void> {
    this.setState({ loading: true, error: null });

    try {
      let cardData: any;

      console.log('🚀🚀🚀 LATEST VERSION 1.0.14.0 - Loading card from:', this.props.cardJsonUrl);
      console.log('🔧🔧🔧 FETCH API FIX ACTIVE - No more "Failed to fetch" errors - NEW BUNDLE! 🔧🔧🔧');
      console.log('💥💥💥 v1.0.34.0 CACHE BREAK - TIMESTAMP: ' + Date.now() + ' 💥💥💥');

      // Check if cardJsonUrl is a predefined template
      if (this.props.cardJsonUrl.startsWith('template:')) {
        const templateName = this.props.cardJsonUrl.replace('template:', '');
        console.log('🎯 FIXED VERSION - Loading template:', templateName);
        cardData = (cardTemplates as any)[templateName];
        if (!cardData) {
          throw new Error(`Template '${templateName}' not found`);
        }
      }
      // Check if cardJsonUrl is a React component
      else if (this.props.cardJsonUrl.startsWith('component:')) {
        const componentName = this.props.cardJsonUrl.replace('component:', '');
        console.log('🎯 COMPONENT MODE - Loading component:', componentName);
        this.renderComponent(componentName);
        return;
      }
      // Check if cardJsonUrl is a static asset
      else if (this.props.cardJsonUrl.startsWith('asset:')) {
        const assetName = this.props.cardJsonUrl.replace('asset:', '');
        console.log('🎯 FIXED VERSION - Loading asset:', assetName);
        try {
          cardData = await this.loadAssetCard(assetName);
          console.log('🎯 FIXED VERSION - Asset loaded successfully:', cardData);
        } catch (assetError) {
          console.error('🎯 FIXED VERSION - Error loading asset:', assetError);
          throw assetError;
        }
      }
      // Load from URL (original behavior) - only for actual HTTP URLs
      else if (this.props.cardJsonUrl.startsWith('http://') || this.props.cardJsonUrl.startsWith('https://')) {
        console.log('Loading from URL:', this.props.cardJsonUrl);
        const response: HttpClientResponse = await this.props.context.httpClient.get(
          this.props.cardJsonUrl,
          HttpClient.configurations.v1
        );

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        cardData = await response.json();
      }
      // Invalid URL scheme
      else {
        throw new Error(`Invalid card source: ${this.props.cardJsonUrl}. Must be template:, asset:, or a valid HTTP(S) URL.`);
      }

      this.setState({ cardData, loading: false });
      this.renderAdaptiveCard(cardData);
    } catch (error) {
      this.setState({ 
        error: error.message || 'Failed to load Adaptive Card',
        loading: false 
      });
    }
  }

  private async loadAssetCard(assetName: string): Promise<any> {
    // Embedded JSON cards for common templates
    const embeddedCards: { [key: string]: any } = {
      'sample-card': {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.3",
        "body": [
          {
            "type": "TextBlock",
            "text": "Sample Adaptive Card",
            "weight": "Bolder",
            "size": "Large",
            "color": "Accent"
          },
          {
            "type": "TextBlock",
            "text": "This card is loaded from embedded JSON in the SPFx solution.",
            "wrap": true,
            "spacing": "Medium"
          }
        ]
      },
      'project-status': {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.3",
        "body": [
          {
            "type": "TextBlock",
            "text": "Project Status",
            "weight": "Bolder",
            "size": "Large",
            "color": "Good"
          },
          {
            "type": "FactSet",
            "facts": [
              {
                "title": "Status:",
                "value": "In Progress"
              },
              {
                "title": "Completion:",
                "value": "75%"
              }
            ]
          }
        ]
      },
      'team-notification': {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.3",
        "body": [
          {
            "type": "TextBlock",
            "text": "🚀 Team Notification",
            "weight": "Bolder",
            "size": "Large",
            "color": "Accent"
          },
          {
            "type": "TextBlock",
            "text": "New feature release available! Check out the enhanced Adaptive Cards integration.",
            "wrap": true,
            "spacing": "Medium"
          },
          {
            "type": "FactSet",
            "facts": [
              {
                "title": "Release Date:",
                "value": "July 30, 2025"
              },
              {
                "title": "Version:",
                "value": "1.0.9.0"
              }
            ]
          }
        ]
      },
      'sales-dashboard': {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.3",
        "body": [
          {
            "type": "TextBlock",
            "text": "📊 Sales Dashboard",
            "weight": "Bolder",
            "size": "Large",
            "color": "Accent"
          },
          {
            "type": "TextBlock",
            "text": "Monthly Revenue: $125,000 (↗️ +15%)",
            "weight": "Bolder",
            "color": "Good",
            "spacing": "Medium"
          },
          {
            "type": "FactSet",
            "facts": [
              {
                "title": "Active Deals:",
                "value": "47"
              },
              {
                "title": "Top Performer:",
                "value": "Sarah Johnson - $45,000"
              }
            ]
          }
        ]
      }
    };

    const card = embeddedCards[assetName];
    if (!card) {
      throw new Error(`Asset '${assetName}' not found`);
    }
    
    // Return the card directly (no fetch needed since it's embedded)
    return Promise.resolve(card);
  }

  private renderAdaptiveCard(cardJson: any): void {
    if (!this.cardContainer.current) return;

    // Clear previous content
    this.cardContainer.current.innerHTML = '';

    try {
      // Create Adaptive Card instance
      const adaptiveCard = new AdaptiveCards.AdaptiveCard();
      
      // Set up action handling
      adaptiveCard.onExecuteAction = (action: AdaptiveCards.Action) => {
        this.handleSubmitAction(action);
      };

      // Parse and render the card
      adaptiveCard.parse(cardJson);
      const renderedCard = adaptiveCard.render();
      
      if (renderedCard) {
        this.cardContainer.current.appendChild(renderedCard);
      }
    } catch (error) {
      console.error('Error rendering Adaptive Card:', error);
      this.setState({ error: 'Failed to render Adaptive Card' });
    }
  }

  private async handleSubmitAction(action: AdaptiveCards.Action): Promise<void> {
    try {
      if (action instanceof AdaptiveCards.SubmitAction) {
        const data = action.data as any; // Cast to any for flexibility with dynamic data
        
        if (data && data.action === 'markAsRead' && data.messageId) {
          // If we have a messageId, mark it as read in SharePoint
          if (typeof data.messageId === 'number') {
            await enhancedDataService.markMessageAsRead(data.messageId);
            console.log(`Message ${data.messageId} marked as read successfully`);
            
            // Show success notification
            this.showSuccessMessage('Message marked as read');
          } else {
            console.log('Sample card action - would mark message as read:', data);
          }
        } else {
          console.log('Submit action data:', data);
          // Handle other types of submit actions here
        }
      } else {
        console.log('Non-submit action executed:', action);
      }
    } catch (error) {
      console.error('Error handling submit action:', error);
      this.showErrorMessage('Failed to process action');
    }
  }

  private showSuccessMessage(message: string): void {
    // Simple success indicator - you could enhance this with a proper notification system
    const successDiv = document.createElement('div');
    successDiv.innerHTML = `✓ ${message}`;
    successDiv.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #107c10;
      color: white;
      padding: 12px 16px;
      border-radius: 4px;
      z-index: 1000;
      font-family: 'Segoe UI', sans-serif;
      box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    `;
    
    document.body.appendChild(successDiv);
    setTimeout(() => {
      if (successDiv.parentNode) {
        successDiv.parentNode.removeChild(successDiv);
      }
    }, 3000);
  }

  private showErrorMessage(message: string): void {
    // Simple error indicator
    const errorDiv = document.createElement('div');
    errorDiv.innerHTML = `⚠ ${message}`;
    errorDiv.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #d13438;
      color: white;
      padding: 12px 16px;
      border-radius: 4px;
      z-index: 1000;
      font-family: 'Segoe UI', sans-serif;
      box-shadow: 0 2px 8px rgba(0,0,0,0.2);
    `;
    
    document.body.appendChild(errorDiv);
    setTimeout(() => {
      if (errorDiv.parentNode) {
        errorDiv.parentNode.removeChild(errorDiv);
      }
    }, 5000);
  }

  private getDefaultCard(): any {
    return {
      "type": "AdaptiveCard",
      "version": "1.5",
      "body": [
        {
          "type": "TextBlock",
          "text": "Welcome to Adaptive Cards!",
          "size": "Large",
          "weight": "Bolder"
        },
        {
          "type": "TextBlock",
          "text": "This is a sample Adaptive Card showing integration with SharePoint Framework. Configure the web part to load your own card JSON.",
          "wrap": true
        },
        {
          "type": "FactSet",
          "facts": [
            {
              "title": "Framework:",
              "value": "SharePoint Framework"
            },
            {
              "title": "Technology:",
              "value": "Adaptive Cards"
            },
            {
              "title": "Version:",
              "value": "1.5"
            },
            {
              "title": "Integration:",
              "value": "SharePoint Lists & Power Automate"
            }
          ]
        }
      ],
      "actions": [
        {
          "type": "Action.OpenUrl",
          "title": "Learn More",
          "url": "https://adaptivecards.io/"
        },
        {
          "type": "Action.Submit",
          "title": "Mark as Read",
          "data": {
            "action": "markAsRead",
            "messageId": "sample"
          }
        }
      ]
    };
  }

  private renderPlaceholder(): React.ReactElement {
    return (
      <div className={styles.placeholder}>
        <div className={styles.placeholderIcon}>📋</div>
        <div className={styles.placeholderTitle}>Configure your Adaptive Card</div>
        <div className={styles.placeholderDescription}>
          Please configure the Card JSON URL in the web part properties.
        </div>
        <button 
          className={styles.configureButton}
          onClick={() => this.props.context.propertyPane.open()}
        >
          Configure
        </button>
      </div>
    );
  }

  private renderTitle(): React.ReactElement | null {
    if (this.props.displayMode === DisplayMode.Edit) {
      return (
        <input
          type="text"
          value={this.props.title}
          onChange={(e) => this.props.updateProperty(e.target.value)}
          placeholder="Enter web part title"
          style={{
            fontSize: '18px',
            fontWeight: 'bold',
            border: '1px dashed #ccc',
            padding: '4px 8px',
            background: 'transparent',
            width: '100%',
            marginBottom: '16px'
          }}
        />
      );
    }

    return this.props.title ? (
      <h2 style={{ marginBottom: '16px', fontSize: '18px', fontWeight: 'bold' }}>
        {escape(this.props.title)}
      </h2>
    ) : null;
  }

  private renderComponent(componentName: string): void {
    this.setState({ loading: false, error: null });
    // Component rendering will happen in the render method
  }

  public render(): React.ReactElement<IAdaptiveCardProps> {
    console.log('🎨 AdaptiveCardComponent.render() called');
    console.log('🎨 Props:', this.props);
    console.log('🎨 State:', this.state);
    console.log('🎨 Card JSON URL:', this.props.cardJsonUrl);
    
    // Check if we're in component mode
    if (this.props.cardJsonUrl?.startsWith('component:')) {
      const componentName = this.props.cardJsonUrl.replace('component:', '');
      console.log('🎨 Component mode detected:', componentName);
      
      switch (componentName) {
        case 'teams-message-creator':
          console.log('🎨 Rendering TeamsMessageCreator component');
          return (
            <div className={styles.adaptiveCardComponent}>
              <TeamsMessageCreator context={this.props.context} />
            </div>
          );
        case 'manager-dashboard':
          console.log('🎨 Rendering ManagerDashboard component');
          // Enhanced data service is globally available
          
          return (
            <div className={styles.adaptiveCardComponent}>
              <ManagerDashboard />
            </div>
          );
        case 'message-list-diagnostic':
          console.log('🎨 MessageListDiagnostic component removed');
          return (
            <div className={styles.adaptiveCardComponent}>
              <div>MessageListDiagnostic component has been removed as part of cleanup</div>
            </div>
          );
        default:
          console.log('🎨 Unknown component, rendering error');
          return (
            <div className={styles.adaptiveCardComponent}>
              <div className={styles.error}>
                <div className={styles.errorIcon}>⚠️</div>
                <div>Unknown component: {componentName}</div>
              </div>
            </div>
          );
      }
    }

    // Only show placeholder in edit mode when no URL is configured AND it's not loading/showing content
    if (this.props.displayMode === DisplayMode.Edit && !this.props.cardJsonUrl && !this.state.cardData && !this.state.loading) {
      console.log('🎨 Rendering placeholder (edit mode, no URL)');
      return (
        <div className={styles.adaptiveCardComponent}>
          {this.renderPlaceholder()}
        </div>
      );
    }

    console.log('🎨 Rendering main component with state-based content');
    return (
      <div className={styles.adaptiveCardComponent}>
        {this.renderTitle()}
        
        {this.state.loading && (
          <div className={styles.loading}>
            <div className={styles.spinner}></div>
            <span>Loading Adaptive Card...</span>
          </div>
        )}

        {this.state.error && (
          <div className={styles.error}>
            <div className={styles.errorIcon}>⚠️</div>
            <div>
              <strong>Error loading Adaptive Card:</strong>
              <br />
              {this.state.error}
              <br />
              <small>URL: {this.props.cardJsonUrl}</small>
            </div>
          </div>
        )}

        {!this.state.loading && !this.state.error && (
          <div ref={this.cardContainer} className={styles.cardContainer}>
            {/* Adaptive card will be rendered here */}
          </div>
        )}
      </div>
    );
  }
}
