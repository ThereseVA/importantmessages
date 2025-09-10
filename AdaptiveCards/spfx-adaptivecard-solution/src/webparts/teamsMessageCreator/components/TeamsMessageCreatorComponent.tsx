import * as React from 'react';
import { ITeamsMessageCreatorProps } from './ITeamsMessageCreatorProps';
import { TeamsMessageCreator } from '../../adaptiveCardViewer/components/TeamsMessageCreator';
import { EnhancedDataService } from '../../../services/EnhancedDataService';

interface ITeamsMessageCreatorComponentState {
  isManager: boolean;
  loading: boolean;
  error: string | null;
}

export class TeamsMessageCreatorComponent extends React.Component<ITeamsMessageCreatorProps, ITeamsMessageCreatorComponentState> {
  private dataService: EnhancedDataService;

  constructor(props: ITeamsMessageCreatorProps) {
    super(props);
    
    this.state = {
      isManager: false,
      loading: true,
      error: null
    };

    this.dataService = new EnhancedDataService();
  }

  public async componentDidMount(): Promise<void> {
    try {
      await this.dataService.initialize(this.props.context);
      
      // Check if user is manager
      const userInfo = await this.dataService.getCurrentUser();
      const isManager = userInfo.isManager || userInfo.isAdmin;
      
      console.log('👤 Teams Message Creator user check:', { 
        displayName: userInfo.displayName, 
        isManager,
        isAdmin: userInfo.isAdmin 
      });

      this.setState({ 
        isManager, 
        loading: false 
      });
    } catch (error) {
      console.error('❌ Error checking user permissions:', error);
      this.setState({ 
        error: error.message || 'Failed to verify permissions',
        loading: false 
      });
    }
  }

  public render(): React.ReactElement<ITeamsMessageCreatorProps> {
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
          <p style={{ margin: 0, color: '#666', fontSize: '14px' }}>
            This Teams Message Creator is only accessible to managers and administrators.
          </p>
          <p style={{ margin: '12px 0 0 0', color: '#666', fontSize: '12px' }}>
            Contact your administrator if you need access to this feature.
          </p>
        </div>
      );
    }

    return (
      <div>
        <div style={{ 
          marginBottom: '20px', 
          padding: '16px', 
          backgroundColor: '#f3f9ff', 
          border: '1px solid #0078d4',
          borderRadius: '4px'
        }}>
          <h2 style={{ margin: '0 0 8px 0', color: '#0078d4' }}>📝 Teams Message Creator</h2>
          <p style={{ margin: 0, color: '#666', fontSize: '14px' }}>
            Create and send messages to Teams channels and user groups.
          </p>
        </div>
        
        <TeamsMessageCreator 
          context={this.props.context}
          dataService={this.dataService}
        />
      </div>
    );
  }
}
