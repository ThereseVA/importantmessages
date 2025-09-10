import * as React from 'react';
import { IManagerDashboardProps } from './IManagerDashboardProps';
import { enhancedDataService } from '../../../services/EnhancedDataService';

interface IManagerDashboardComponentState {
  isManager: boolean;
  loading: boolean;
  error: string | null;
}

export class ManagerDashboardComponent extends React.Component<IManagerDashboardProps, IManagerDashboardComponentState> {
  constructor(props: IManagerDashboardProps) {
    super(props);
    this.state = {
      isManager: false,
      loading: true,
      error: null
    };
  }

  public async componentDidMount(): Promise<void> {
    try {
      // Initialize the enhanced data service with the current context
      if (this.props.context) {
        await enhancedDataService.initialize(this.props.context);
      }
      
      // Check if user is a manager using the SharePoint Managers list
      const isManager = await enhancedDataService.isCurrentUserManager();
      
      console.log('ManagerDashboard: Manager status from SharePoint list:', isManager);
      
      this.setState({
        isManager,
        loading: false,
        error: null
      });
      
    } catch (error) {
      console.error('Error checking manager status from SharePoint list:', error);
      this.setState({
        isManager: false,
        loading: false,
        error: 'Failed to verify manager access from SharePoint Managers list'
      });
    }
  }

  public render(): React.ReactElement<IManagerDashboardProps> {
    const { loading, isManager, error } = this.state;

    if (loading) {
      return (
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <div style={{ fontSize: '16px', color: '#666' }}>🔄 Checking permissions...</div>
        </div>
      );
    }

    if (error) {
      return (
        <div style={{ padding: '20px', border: '1px solid #d13438', borderRadius: '4px', backgroundColor: '#fdf2f2' }}>
          <h3 style={{ color: '#d13438', margin: '0 0 10px 0' }}>⚠️ Access Error</h3>
          <p style={{ margin: 0, color: '#666' }}>{error}</p>
        </div>
      );
    }

    if (!isManager) {
      return (
        <div style={{ padding: '30px', textAlign: 'center', border: '1px solid #ffbe00', borderRadius: '8px', backgroundColor: '#fffbf0' }}>
          <div style={{ fontSize: '48px', marginBottom: '16px' }}>🔒</div>
          <h3 style={{ color: '#d83b01', margin: '0 0 12px 0' }}>Manager Access Required</h3>
          <p style={{ margin: '0 0 12px 0', color: '#666', fontSize: '14px' }}>
            This Manager Dashboard is only accessible to managers listed in the SharePoint Managers list.
          </p>
          <div style={{ 
            padding: '12px', 
            backgroundColor: '#fff3cd', 
            borderRadius: '4px', 
            textAlign: 'left',
            fontSize: '12px',
            color: '#856404'
          }}>
            <strong>How manager access is determined:</strong>
            <ul style={{ margin: '8px 0 0 0', paddingLeft: '20px' }}>
              <li>Your email must be listed in the "Managers" SharePoint list</li>
              <li>Your entry must have "Is Active" set to "Yes"</li>
              <li>Contact HR or IT to be added to the managers list</li>
            </ul>
          </div>
          <p style={{ margin: '12px 0 0 0', color: '#666', fontSize: '12px' }}>
            Contact your administrator if you believe you should have manager access.
          </p>
        </div>
      );
    }

    return (
      <div>
        <div style={{ 
          marginBottom: '20px', 
          padding: '16px', 
          backgroundColor: '#f0f8ff', 
          border: '1px solid #107c10',
          borderRadius: '4px'
        }}>
          <h2 style={{ margin: '0 0 8px 0', color: '#107c10' }}>🎛️ Manager Dashboard</h2>
          <p style={{ margin: 0, color: '#666', fontSize: '14px' }}>
            Comprehensive management tools for messages, analytics, and system administration.
          </p>
        </div>
        
        <div style={{ padding: '20px', border: '1px solid #ddd', borderRadius: '8px' }}>
          <h3>📊 Management Features</h3>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginTop: '16px' }}>
            <div style={{ padding: '16px', border: '1px solid #107c10', borderRadius: '4px', backgroundColor: '#f8f9fa' }}>
              <h4 style={{ margin: '0 0 8px 0', color: '#107c10' }}>📈 Message Analytics</h4>
              <p style={{ margin: 0, fontSize: '14px', color: '#666' }}>
                View comprehensive statistics about message delivery and engagement.
              </p>
            </div>
            <div style={{ padding: '16px', border: '1px solid #0078d4', borderRadius: '4px', backgroundColor: '#f8f9fa' }}>
              <h4 style={{ margin: '0 0 8px 0', color: '#0078d4' }}>📝 Message Creation</h4>
              <p style={{ margin: 0, fontSize: '14px', color: '#666' }}>
                Create and send new messages to specific teams or user groups.
              </p>
            </div>
            <div style={{ padding: '16px', border: '1px solid #d83b01', borderRadius: '4px', backgroundColor: '#f8f9fa' }}>
              <h4 style={{ margin: '0 0 8px 0', color: '#d83b01' }}>⚙️ System Settings</h4>
              <p style={{ margin: 0, fontSize: '14px', color: '#666' }}>
                Configure system-wide settings and manage user permissions.
              </p>
            </div>
            <div style={{ padding: '16px', border: '1px solid #8b5cf6', borderRadius: '4px', backgroundColor: '#f8f9fa' }}>
              <h4 style={{ margin: '0 0 8px 0', color: '#8b5cf6' }}>👥 User Management</h4>
              <p style={{ margin: 0, fontSize: '14px', color: '#666' }}>
                Manage user roles, groups, and access permissions.
              </p>
            </div>
          </div>
          
          <div style={{ marginTop: '24px', padding: '16px', backgroundColor: '#fff3cd', border: '1px solid #ffbe00', borderRadius: '4px' }}>
            <h4 style={{ margin: '0 0 8px 0', color: '#856404' }}>🚧 Advanced Features Coming Soon</h4>
            <p style={{ margin: 0, fontSize: '14px', color: '#856404' }}>
              More advanced management features are in development. This dashboard will be expanded with additional capabilities.
            </p>
          </div>
          
          <div style={{ marginTop: '16px', textAlign: 'center' }}>
            <p style={{ fontSize: '12px', color: '#666' }}>
              For access to the full Manager Dashboard functionality, use the Adaptive Card Viewer with <code>?cardSource=manager-dashboard</code> URL parameter.
            </p>
          </div>
        </div>
      </div>
    );
  }
}
