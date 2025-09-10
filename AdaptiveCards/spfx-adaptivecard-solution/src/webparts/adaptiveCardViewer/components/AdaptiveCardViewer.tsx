import * as React from 'react';
import { 
  Stack, 
  PrimaryButton, 
  DefaultButton, 
  Spinner, 
  SpinnerSize,
  Text,
  Icon
} from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { EmployeeMessageList } from './EmployeeMessageList';
import { ManagerDashboard } from './ManagerDashboard';
import { TeamsMessageCreator } from './TeamsMessageCreator';
import { AdaptiveCardComponent } from './AdaptiveCardComponent';
import styles from './AdaptiveCardViewer.module.scss';

// Missing interface definitions
export interface IAdaptiveCardViewerProps {
  context: WebPartContext;
  messageId?: number;
  description?: string;
}

export interface IAdaptiveCardViewerState {
  isLoading: boolean;
  currentView: 'messages' | 'dashboard' | 'creator' | 'diagnostic';
  userRole: 'manager' | 'employee';
  userRoleDetails?: {
    role: 'Employee' | 'Manager' | 'Admin' | 'SuperAdmin';
    method: string;
    isManager: boolean;
  };
  error?: string;
}

/**
 * Adaptive Card Viewer Component
 * Displays adaptive cards based on user role and selected view
 */
export default class AdaptiveCardViewer extends React.Component<IAdaptiveCardViewerProps, IAdaptiveCardViewerState> {

  constructor(props: IAdaptiveCardViewerProps) {
    super(props);

    this.state = {
      isLoading: true,
      currentView: 'messages',
      userRole: 'employee',
      error: undefined
    };

    // Enhanced services are initialized globally
  }

  public async componentDidMount(): Promise<void> {
    try {
      console.log('🔄 AdaptiveCardViewer: Component mounting...');
      
      // Initialize enhanced data service
      await enhancedDataService.initialize(this.props.context, this.props.context?.pageContext?.web?.absoluteUrl);
      
      // Get current user to determine role (simplified - assume employee for now)
      const currentUser = enhancedDataService.getCurrentUser();
      const userRole = 'employee'; // TODO: Implement role detection in enhanced service
      
      this.setState({ 
        userRole: userRole,
        isLoading: false,
        error: undefined 
      });
      
      console.log(`🎯 User role determined: ${userRole}`);
      
    } catch (error) {
      console.error('❌ Error during component mount:', error);
      this.setState({ 
        error: `Kunde inte ladda komponenten: ${error.message}`,
        isLoading: false,
        userRole: 'employee' // Safe fallback
      });
    }
  }

  public render(): React.ReactElement<IAdaptiveCardViewerProps> {
    const { isLoading, currentView, userRole, error } = this.state;

    if (isLoading) {
      return (
        <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
          <Spinner size={SpinnerSize.large} label="Laddar..." />
        </Stack>
      );
    }

    if (error) {
      return (
        <div className={styles.adaptiveCardViewer}>
          <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
            <Icon iconName="Error" style={{ fontSize: '48px', color: '#d13438', marginBottom: '16px' }} />
            <Text variant="large" style={{ color: '#d13438', marginBottom: '8px' }}>
              Ett fel uppstod
            </Text>
            <Text variant="medium" style={{ color: '#605e5c', textAlign: 'center' }}>
              {error}
            </Text>
            <PrimaryButton
              text="Försök igen"
              iconProps={{ iconName: 'Refresh' }}
              onClick={() => window.location.reload()}
              style={{ marginTop: '16px' }}
            />
          </Stack>
        </div>
      );
    }

    return (
      <div className={styles.adaptiveCardViewer}>
        {/* User role indicator for debugging */}
        <Stack style={{ marginBottom: 10 }}>
          <Text variant="small" style={{ color: '#666' }}>
            Inloggad som: {userRole === 'manager' ? '👑 Chef' : '👤 Medarbetare'} ({this.props.context.pageContext.user.email})
          </Text>
        </Stack>

        {/* Navigation based on user role */}
        {userRole === 'manager' ? (
          <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginBottom: 20 }}>
            <PrimaryButton
              text="📊 Manager Dashboard"
              onClick={() => this.setState({ currentView: 'dashboard', error: undefined })}
              disabled={currentView === 'dashboard'}
            />
            <DefaultButton
              text="✉️ Skapa meddelande"
              onClick={() => this.setState({ currentView: 'creator', error: undefined })}
              disabled={currentView === 'creator'}
            />
            <DefaultButton
              text="📬 Mina meddelanden"
              onClick={() => this.setState({ currentView: 'messages', error: undefined })}
              disabled={currentView === 'messages'}
            />
            <DefaultButton
              text="🔍 Diagnostik"
              onClick={() => this.setState({ currentView: 'diagnostic', error: undefined })}
              disabled={currentView === 'diagnostic'}
            />
          </Stack>
        ) : (
          <Stack horizontal tokens={{ childrenGap: 16 }} style={{ marginBottom: 20 }}>
            <PrimaryButton
              text="📬 Mina meddelanden"
              onClick={() => this.setState({ currentView: 'messages', error: undefined })}
              disabled={currentView === 'messages'}
            />
          </Stack>
        )}

        {/* Content based on current view and user role */}
        {this.renderCurrentView()}
      </div>
    );
  }

  private renderCurrentView(): React.ReactElement {
    const { currentView, userRole } = this.state;

    try {
      switch (currentView) {
        case 'messages':
          return (
            <EmployeeMessageList />
          );

        case 'dashboard':
          if (userRole !== 'manager') {
            return (
              <Stack horizontalAlign="center" style={{ padding: '20px' }}>
                <Text variant="large" style={{ color: '#d13438' }}>
                  ❌ Du har inte behörighet att se denna sida
                </Text>
              </Stack>
            );
          }
          return (
            <ManagerDashboard />
          );

        case 'creator':
          if (userRole !== 'manager') {
            return (
              <Stack horizontalAlign="center" style={{ padding: '20px' }}>
                <Text variant="large" style={{ color: '#d13438' }}>
                  ❌ Du har inte behörighet att skapa meddelanden
                </Text>
              </Stack>
            );
          }
          return (
            <TeamsMessageCreator
              context={this.props.context}
              onMessageCreated={() => {
                console.log('Message created, refreshing dashboard...');
                this.setState({ currentView: 'dashboard' });
              }}
            />
          );

        case 'diagnostic':
          if (userRole !== 'manager') {
            return (
              <Stack horizontalAlign="center" style={{ padding: '20px' }}>
                <Text variant="large" style={{ color: '#d13438' }}>
                  ❌ Du har inte behörighet att se diagnostik
                </Text>
              </Stack>
            );
          }
          return (
            <Stack>
              <Text variant="large">Diagnostic tools have been replaced with enhanced services.</Text>
              <Text>Use the Messages or Manager Dashboard views instead.</Text>
            </Stack>
          );

        default:
          return (
            <Stack horizontalAlign="center" style={{ padding: '20px' }}>
              <Text variant="large" style={{ color: '#d13438' }}>
                ❌ Okänd vy: {currentView}
              </Text>
            </Stack>
          );
      }
    } catch (error) {
      console.error('❌ Error rendering view:', error);
      return (
        <Stack horizontalAlign="center" style={{ padding: '20px' }}>
          <Text variant="large" style={{ color: '#d13438' }}>
            ❌ Ett fel uppstod när sidan skulle laddas
          </Text>
          <Text variant="medium" style={{ color: '#666' }}>
            {error.message}
          </Text>
        </Stack>
      );
    }
  }
}