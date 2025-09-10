import * as React from 'react';
import styles from './DashboardComponent.module.scss';
import { IDashboardProps } from './IDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';
import { enhancedDataService, IMessage } from '../../../services/EnhancedDataService';

export interface IDashboardState {
  messages: IMessage[];
  filteredMessages: IMessage[];
  loading: boolean;
  error: string | null;
  lastRefresh: Date | null;
  customSiteUrl: string;
  filters: {
    priority: string;
    readStatus: string;
    targetAudience: string;
    dateRange: string;
  };
  showCharts: boolean;
}

export class DashboardComponent extends React.Component<IDashboardProps, IDashboardState> {
  private refreshTimer: number | null = null;

  constructor(props: IDashboardProps) {
    super(props);
    this.state = {
      messages: [],
      filteredMessages: [],
      loading: false,
      error: null,
      lastRefresh: null,
      customSiteUrl: '',
      filters: {
        priority: 'All',
        readStatus: 'All',
        targetAudience: 'All',
        dateRange: 'All'
      },
      showCharts: false
    };
  }

  public async componentDidMount(): Promise<void> {
    // Initialize the enhanced data service with dataSourceUrl
    await enhancedDataService.initialize(this.props.context, this.props.dataSourceUrl);
    
    // Check if we're in Teams context and handle accordingly
    this.handleTeamsContext();
    
    // Load initial data
    this.loadMessages();

    // Set up auto-refresh if configured
    if (this.props.refreshInterval > 0) {
      this.refreshTimer = setInterval(() => {
        this.loadMessages();
      }, this.props.refreshInterval);
    }
  }

  private handleTeamsContext(): void {
    try {
      const url = window.location.href;
      const isTeamsUrl = url.includes('teams.microsoft.com') || url.includes('teams.office.com');
      const hasTeamsContext = this.props.context.sdks?.microsoftTeams?.context !== undefined;
      
      console.log('Dashboard: Teams context check:', {
        currentUrl: url,
        isTeamsUrl,
        hasTeamsContext,
        dataSourceUrl: this.props.dataSourceUrl
      });
      
      if (isTeamsUrl || hasTeamsContext) {
        console.log('Dashboard: Running in Teams context');
        
        // If a dataSourceUrl is configured, extract the SharePoint site URL from it
        if (this.props.dataSourceUrl && this.props.dataSourceUrl.includes('sharepoint.com')) {
          const match = this.props.dataSourceUrl.match(/(https:\/\/[^\/]+\/[^\/]+\/[^\/]+)/);
          if (match) {
            const sharePointSite = match[1];
            console.log('Dashboard: Setting SharePoint site for Teams:', sharePointSite);
            enhancedDataService.setSharePointSiteUrl(sharePointSite);
            this.setState({ customSiteUrl: sharePointSite });
          }
        } else {
          // Try to get SharePoint site from Teams context
          if (this.props.context.sdks?.microsoftTeams?.context) {
            const teamsContext = this.props.context.sdks.microsoftTeams.context;
            let sharePointSite = null;
            
            if (teamsContext.sharepoint?.webAbsoluteUrl) {
              sharePointSite = teamsContext.sharepoint.webAbsoluteUrl;
            } else if (teamsContext.teamSiteUrl) {
              sharePointSite = teamsContext.teamSiteUrl;
            }
            
            if (sharePointSite) {
              console.log('Dashboard: Using SharePoint site from Teams context:', sharePointSite);
              enhancedDataService.setSharePointSiteUrl(sharePointSite);
              this.setState({ customSiteUrl: sharePointSite });
            } else {
              console.warn('Dashboard: No SharePoint site found in Teams context. Please configure dataSourceUrl in web part properties.');
            }
          }
        }
      } else {
        console.log('Dashboard: Running in SharePoint context');
      }
    } catch (error) {
      console.error('Dashboard: Error handling Teams context:', error);
    }
  }

  public componentWillUnmount(): void {
    if (this.refreshTimer) {
      clearInterval(this.refreshTimer);
    }
  }

  public componentDidUpdate(prevProps: IDashboardProps): void {
    // Restart timer if refresh interval changed
    if (prevProps.refreshInterval !== this.props.refreshInterval) {
      if (this.refreshTimer) {
        clearInterval(this.refreshTimer);
      }

      if (this.props.refreshInterval > 0) {
        this.refreshTimer = setInterval(() => {
          this.loadMessages();
        }, this.props.refreshInterval);
      }
    }
  }

  private isTeamsContext(): boolean {
    const url = window.location.href;
    const isTeamsUrl = url.includes('teams.microsoft.com') || url.includes('teams.office.com');
    const hasTeamsContext = this.props.context.sdks?.microsoftTeams?.context !== undefined;
    return isTeamsUrl || hasTeamsContext;
  }

  private async loadMessages(): Promise<void> {
    this.setState({ loading: true, error: null });

    try {
      console.log('Dashboard: Starting to load messages...');
      console.log('Dashboard: Current site URL:', this.props.context?.pageContext?.web?.absoluteUrl);
      console.log('Dashboard: DataService custom site URL:', this.state.customSiteUrl);
      
      const messages = await enhancedDataService.getMessagesForCurrentUser();
      console.log('Dashboard: Successfully loaded messages:', messages.length);
      
      const filteredMessages = this.applyFilters(messages);
      this.setState({ 
        messages, 
        filteredMessages,
        loading: false, 
        lastRefresh: new Date() 
      });
    } catch (error) {
      console.error('Dashboard: Error loading messages:', error);
      console.error('Dashboard: Error details:', {
        message: error.message,
        stack: error.stack,
        name: error.name
      });
      
      this.setState({ 
        error: `${error.message || 'Failed to load messages'} (Check browser console for details)`,
        loading: false 
      });
    }
  }

  private async handleMarkAsRead(messageId: number): Promise<void> {
    try {
      await enhancedDataService.markMessageAsRead(messageId);
      
      // Update the local state to reflect the read status
      this.setState(prevState => ({
        messages: prevState.messages.map(msg => 
          msg.Id === messageId 
            ? { ...msg, ReadBy: (msg.ReadBy || '') + ';' + this.props.context.pageContext.user.email }
            : msg
        )
      }));
    } catch (error) {
      console.error('Error marking message as read:', error);
      // You could show a toast notification here
    }
  }

  private isMessageRead(message: IMessage): boolean {
    const currentUserEmail = this.props.context.pageContext.user.email;
    return message.ReadBy?.includes(currentUserEmail) || false;
  }

  private getPriorityColor(priority: string): string {
    switch (priority) {
      case 'High': return '#d13438';
      case 'Medium': return '#ff8c00';
      case 'Low': return '#107c10';
      default: return '#605e5c';
    }
  }

  private renderTitle(): React.ReactElement | null {
    if (this.props.displayMode === DisplayMode.Edit) {
      return (
        <input
          type="text"
          value={this.props.title}
          onChange={(e) => this.props.updateProperty(e.target.value)}
          placeholder="Enter dashboard title"
          style={{
            fontSize: '24px',
            fontWeight: 'bold',
            border: '1px dashed #ccc',
            padding: '8px 12px',
            background: 'transparent',
            width: '100%',
            marginBottom: '16px'
          }}
        />
      );
    }

    return this.props.title ? (
      <h1 style={{ marginBottom: '16px', fontSize: '24px', fontWeight: 'bold' }}>
        {escape(this.props.title)}
      </h1>
    ) : null;
  }

  private renderDescription(): React.ReactElement | null {
    if (!this.props.description) return null;

    if (this.props.displayMode === DisplayMode.Edit) {
      return (
        <textarea
          value={this.props.description}
          onChange={(e) => this.props.updateProperty(e.target.value)}
          placeholder="Enter dashboard description"
          style={{
            fontSize: '14px',
            border: '1px dashed #ccc',
            padding: '8px 12px',
            background: 'transparent',
            width: '100%',
            marginBottom: '16px',
            minHeight: '60px',
            resize: 'vertical'
          }}
        />
      );
    }

    return (
      <p style={{ marginBottom: '16px', color: '#605e5c' }}>
        {escape(this.props.description)}
      </p>
    );
  }

  private renderRefreshInfo(): React.ReactElement | null {
    if (!this.state.lastRefresh) return null;

    return (
      <div className={styles.refreshInfo}>
        <span>Last updated: {this.state.lastRefresh.toLocaleTimeString()}</span>
        {this.props.showRefreshButton && (
          <button 
            className={styles.refreshButton}
            onClick={() => this.loadMessages()}
            disabled={this.state.loading}
          >
            {this.state.loading ? '⟳' : '🔄'} Refresh
          </button>
        )}
      </div>
    );
  }

  private renderMessage(message: IMessage): React.ReactElement {
    const isRead = this.isMessageRead(message);
    const priorityColor = this.getPriorityColor(message.Priority);

    return (
      <div 
        key={message.Id} 
        className={`${styles.messageCard} ${isRead ? styles.readMessage : styles.unreadMessage}`}
      >
        <div className={styles.messageHeader}>
          <div className={styles.messageTitle}>
            <div 
              className={styles.priorityIndicator}
              style={{ backgroundColor: priorityColor }}
            ></div>
            <h3>{escape(message.Title)}</h3>
          </div>
          <div className={styles.messageMetadata}>
            <span className={styles.priority} style={{ color: priorityColor }}>
              {message.Priority}
            </span>
            <span className={styles.date}>
              {new Date(message.Created).toLocaleDateString()}
            </span>
          </div>
        </div>

        <div className={styles.messageContent}>
          <div dangerouslySetInnerHTML={{ __html: message.MessageContent }} />
        </div>

        <div className={styles.messageFooter}>
          <div className={styles.messageInfo}>
            <span>From: {escape(message.Author.Title)}</span>
            <span>Expires: {new Date(message.ExpiryDate).toLocaleDateString()}</span>
          </div>
          
          {!isRead && (
            <button 
              className={styles.markReadButton}
              onClick={() => this.handleMarkAsRead(message.Id)}
            >
              Mark as Read
            </button>
          )}
          
          {isRead && (
            <span className={styles.readIndicator}>✓ Read</span>
          )}
        </div>
      </div>
    );
  }

  private renderPlaceholder(): React.ReactElement {
    return (
      <div className={styles.placeholder}>
        <div className={styles.placeholderIcon}>📊</div>
        <div className={styles.placeholderTitle}>Configure your Dashboard</div>
        <div className={styles.placeholderDescription}>
          Please configure the dashboard settings in the web part properties.
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

  // Filter and Chart Methods
  private applyFilters(messages: IMessage[]): IMessage[] {
    return messages.filter(message => {
      // Priority filter
      if (this.state.filters.priority !== 'All' && message.Priority !== this.state.filters.priority) {
        return false;
      }

      // Read status filter
      if (this.state.filters.readStatus !== 'All') {
        const isRead = this.isMessageRead(message);
        if (this.state.filters.readStatus === 'Read' && !isRead) return false;
        if (this.state.filters.readStatus === 'Unread' && isRead) return false;
      }

      // Target audience filter
      if (this.state.filters.targetAudience !== 'All' && message.TargetAudience !== this.state.filters.targetAudience) {
        return false;
      }

      // Date range filter
      if (this.state.filters.dateRange !== 'All') {
        const now = new Date();
        const messageDate = new Date(message.Created);
        const daysDiff = Math.floor((now.getTime() - messageDate.getTime()) / (1000 * 3600 * 24));

        switch (this.state.filters.dateRange) {
          case 'Today':
            if (daysDiff > 0) return false;
            break;
          case 'This Week':
            if (daysDiff > 7) return false;
            break;
          case 'This Month':
            if (daysDiff > 30) return false;
            break;
        }
      }

      return true;
    });
  }

  private handleFilterChange = (filterType: keyof IDashboardState['filters'], value: string): void => {
    this.setState(prevState => {
      const newFilters = { ...prevState.filters, [filterType]: value };
      const filteredMessages = this.applyFilters(prevState.messages);
      return {
        filters: newFilters,
        filteredMessages
      };
    });
  }

  private toggleCharts = (): void => {
    this.setState(prevState => ({ showCharts: !prevState.showCharts }));
  }

  private getChartData() {
    const { filteredMessages } = this.state;
    
    // Priority distribution
    const priorityData = {
      High: filteredMessages.filter(m => m.Priority === 'High').length,
      Medium: filteredMessages.filter(m => m.Priority === 'Medium').length,
      Low: filteredMessages.filter(m => m.Priority === 'Low').length
    };

    // Read status distribution
    const readData = {
      Read: filteredMessages.filter(m => this.isMessageRead(m)).length,
      Unread: filteredMessages.filter(m => !this.isMessageRead(m)).length
    };

    // Messages over time (last 7 days)
    const timeData = [];
    for (let i = 6; i >= 0; i--) {
      const date = new Date();
      date.setDate(date.getDate() - i);
      const dateStr = date.toLocaleDateString();
      const count = filteredMessages.filter(m => {
        const msgDate = new Date(m.Created);
        return msgDate.toLocaleDateString() === dateStr;
      }).length;
      timeData.push({ date: dateStr, count });
    }

    return { priorityData, readData, timeData };
  }

  private renderFilters(): React.ReactElement {
    const { filters } = this.state;
    const audiences = this.state.messages.map((m: any) => m.TargetAudience).filter((value: any, index: number, self: any[]) => self.indexOf(value) === index);
    const uniqueAudiences = ['All', ...audiences];

    return (
      <div style={{ 
        display: 'flex', 
        gap: '16px', 
        marginBottom: '20px', 
        flexWrap: 'wrap',
        alignItems: 'center'
      }}>
        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <label style={{ fontSize: '12px', fontWeight: '600', marginBottom: '4px' }}>Priority</label>
          <select 
            value={filters.priority} 
            onChange={(e) => this.handleFilterChange('priority', e.target.value)}
            style={{ padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' }}
          >
            <option value="All">All Priorities</option>
            <option value="High">High</option>
            <option value="Medium">Medium</option>
            <option value="Low">Low</option>
          </select>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <label style={{ fontSize: '12px', fontWeight: '600', marginBottom: '4px' }}>Status</label>
          <select 
            value={filters.readStatus} 
            onChange={(e) => this.handleFilterChange('readStatus', e.target.value)}
            style={{ padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' }}
          >
            <option value="All">All Messages</option>
            <option value="Read">Read</option>
            <option value="Unread">Unread</option>
          </select>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <label style={{ fontSize: '12px', fontWeight: '600', marginBottom: '4px' }}>Audience</label>
          <select 
            value={filters.targetAudience} 
            onChange={(e) => this.handleFilterChange('targetAudience', e.target.value)}
            style={{ padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' }}
          >
            {uniqueAudiences.map((audience: string) => (
              <option key={audience} value={audience}>{audience}</option>
            ))}
          </select>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <label style={{ fontSize: '12px', fontWeight: '600', marginBottom: '4px' }}>Date Range</label>
          <select 
            value={filters.dateRange} 
            onChange={(e) => this.handleFilterChange('dateRange', e.target.value)}
            style={{ padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' }}
          >
            <option value="All">All Time</option>
            <option value="Today">Today</option>
            <option value="This Week">This Week</option>
            <option value="This Month">This Month</option>
          </select>
        </div>

        <button
          onClick={this.toggleCharts}
          style={{
            padding: '8px 16px',
            background: this.state.showCharts ? '#106ebe' : '#0078d4',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: 'pointer',
            fontSize: '12px',
            fontWeight: '600',
            marginTop: '18px'
          }}
        >
          {this.state.showCharts ? '📊 Hide Charts' : '📈 Show Charts'}
        </button>
      </div>
    );
  }

  private renderCharts(): React.ReactElement | null {
    if (!this.state.showCharts) return null;

    const { priorityData, readData, timeData } = this.getChartData();

    return (
      <div style={{ marginBottom: '20px' }}>
        <div style={{ 
          display: 'grid', 
          gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', 
          gap: '20px',
          marginBottom: '20px'
        }}>
          {/* Priority Chart */}
          <div style={{ 
            background: 'white', 
            padding: '20px', 
            borderRadius: '8px', 
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)' 
          }}>
            <h3 style={{ marginTop: 0, marginBottom: '16px', fontSize: '16px' }}>Priority Distribution</h3>
            <div style={{ height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'space-around' }}>
              <div style={{ textAlign: 'center' }}>
                <div style={{ 
                  width: '60px', 
                  height: '60px', 
                  background: '#d13438', 
                  borderRadius: '50%', 
                  display: 'flex', 
                  alignItems: 'center', 
                  justifyContent: 'center',
                  color: 'white',
                  fontSize: '18px',
                  fontWeight: 'bold',
                  margin: '0 auto 8px auto'
                }}>
                  {priorityData.High}
                </div>
                <div style={{ fontSize: '12px' }}>High</div>
              </div>
              <div style={{ textAlign: 'center' }}>
                <div style={{ 
                  width: '60px', 
                  height: '60px', 
                  background: '#ff8c00', 
                  borderRadius: '50%', 
                  display: 'flex', 
                  alignItems: 'center', 
                  justifyContent: 'center',
                  color: 'white',
                  fontSize: '18px',
                  fontWeight: 'bold',
                  margin: '0 auto 8px auto'
                }}>
                  {priorityData.Medium}
                </div>
                <div style={{ fontSize: '12px' }}>Medium</div>
              </div>
              <div style={{ textAlign: 'center' }}>
                <div style={{ 
                  width: '60px', 
                  height: '60px', 
                  background: '#107c10', 
                  borderRadius: '50%', 
                  display: 'flex', 
                  alignItems: 'center', 
                  justifyContent: 'center',
                  color: 'white',
                  fontSize: '18px',
                  fontWeight: 'bold',
                  margin: '0 auto 8px auto'
                }}>
                  {priorityData.Low}
                </div>
                <div style={{ fontSize: '12px' }}>Low</div>
              </div>
            </div>
          </div>

          {/* Read Status Chart */}
          <div style={{ 
            background: 'white', 
            padding: '20px', 
            borderRadius: '8px', 
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)' 
          }}>
            <h3 style={{ marginTop: 0, marginBottom: '16px', fontSize: '16px' }}>Read Status</h3>
            <div style={{ height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'space-around' }}>
              <div style={{ textAlign: 'center' }}>
                <div style={{ 
                  width: '80px', 
                  height: `${Math.max(20, (readData.Read / Math.max(readData.Read + readData.Unread, 1)) * 150)}px`, 
                  background: '#107c10', 
                  margin: '0 auto 8px auto',
                  display: 'flex',
                  alignItems: 'flex-end',
                  justifyContent: 'center',
                  color: 'white',
                  fontWeight: 'bold',
                  paddingBottom: '8px'
                }}>
                  {readData.Read}
                </div>
                <div style={{ fontSize: '12px' }}>Read</div>
              </div>
              <div style={{ textAlign: 'center' }}>
                <div style={{ 
                  width: '80px', 
                  height: `${Math.max(20, (readData.Unread / Math.max(readData.Read + readData.Unread, 1)) * 150)}px`, 
                  background: '#d13438', 
                  margin: '0 auto 8px auto',
                  display: 'flex',
                  alignItems: 'flex-end',
                  justifyContent: 'center',
                  color: 'white',
                  fontWeight: 'bold',
                  paddingBottom: '8px'
                }}>
                  {readData.Unread}
                </div>
                <div style={{ fontSize: '12px' }}>Unread</div>
              </div>
            </div>
          </div>
        </div>

        {/* Messages Over Time */}
        <div style={{ 
          background: 'white', 
          padding: '20px', 
          borderRadius: '8px', 
          boxShadow: '0 2px 4px rgba(0,0,0,0.1)' 
        }}>
          <h3 style={{ marginTop: 0, marginBottom: '16px', fontSize: '16px' }}>Messages Over Time (Last 7 Days)</h3>
          <div style={{ 
            height: '150px', 
            display: 'flex', 
            alignItems: 'flex-end', 
            justifyContent: 'space-between',
            borderBottom: '1px solid #ccc',
            paddingBottom: '10px'
          }}>
            {timeData.map((day, index) => (
              <div key={index} style={{ textAlign: 'center', flex: 1 }}>
                <div style={{ 
                  height: `${Math.max(10, (day.count / Math.max(...timeData.map(d => d.count), 1)) * 100)}px`, 
                  background: '#0078d4', 
                  margin: '0 auto 8px auto',
                  width: '30px',
                  display: 'flex',
                  alignItems: 'flex-end',
                  justifyContent: 'center',
                  color: 'white',
                  fontSize: '12px',
                  fontWeight: 'bold',
                  paddingBottom: '4px'
                }}>
                  {day.count}
                </div>
                <div style={{ fontSize: '10px', transform: 'rotate(-45deg)', transformOrigin: 'center' }}>
                  {day.date.split('/').slice(0, 2).join('/')}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // Navigation methods for Quick Actions
  private openTeamsMessageCreator = (): void => {
    const baseUrl = window.location.origin + window.location.pathname;
    const newUrl = `${baseUrl}?component=teams-message-creator`;
    window.open(newUrl, '_blank');
  }

  private openManagerDashboard = (): void => {
    const baseUrl = window.location.origin + window.location.pathname;
    const newUrl = `${baseUrl}?component=manager-dashboard`;
    window.open(newUrl, '_blank');
  }

  private openMessageDiagnostics = (): void => {
    const baseUrl = window.location.origin + window.location.pathname;
    const newUrl = `${baseUrl}?component=message-list-diagnostic`;
    window.open(newUrl, '_blank');
  }

  public render(): React.ReactElement<IDashboardProps> {
    const { loading, error, messages, filteredMessages, showCharts } = this.state;

    return (
      <div className={styles.dashboardComponent}>
        <div style={{ marginBottom: '20px' }}>
          <h2>📊 Personal Dashboard</h2>
          <p>Monitor your personalized messages and activity</p>
        </div>

        {/* Quick Actions */}
        <div style={{ marginBottom: '20px', padding: '16px', background: '#e8f4fd', borderRadius: '8px', border: '1px solid #0078d4' }}>
          <h3 style={{ margin: '0 0 12px 0', fontSize: '16px', color: '#323130' }}>🚀 Quick Actions</h3>
          <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
            <button
              onClick={() => this.openTeamsMessageCreator()}
              style={{
                padding: '8px 16px',
                backgroundColor: '#0078d4',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600'
              }}
            >
              📝 Create New Message
            </button>
            <button
              onClick={() => this.openManagerDashboard()}
              style={{
                padding: '8px 16px',
                backgroundColor: '#107c10',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600'
              }}
            >
              👥 Manager Dashboard
            </button>
            <button
              onClick={() => this.openMessageDiagnostics()}
              style={{
                padding: '8px 16px',
                backgroundColor: '#ca5010',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600'
              }}
            >
              🔍 Message Diagnostics
            </button>
          </div>
        </div>
        
        {/* Data Source URL Configuration */}
        <div style={{ marginBottom: '20px', padding: '16px', background: '#f8f9fa', borderRadius: '8px', border: '1px solid #e1e5e9' }}>
          <div style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
            <h3 style={{ margin: 0, fontSize: '16px', color: '#323130' }}>⚙️ Data Source Configuration</h3>
          </div>
          <div style={{ marginBottom: '12px' }}>
            <label style={{ display: 'block', fontSize: '14px', fontWeight: '600', marginBottom: '4px' }}>
              SharePoint Site URL:
            </label>
            <input
              type="text"
              value={this.state.customSiteUrl || ''}
              onChange={(e) => this.setState({ customSiteUrl: e.target.value })}
              placeholder="https://yourtenant.sharepoint.com/sites/yoursite (leave empty to use current site)"
              style={{ 
                width: '100%', 
                padding: '8px 12px', 
                border: '1px solid #d1d1d1', 
                borderRadius: '4px',
                fontSize: '14px'
              }}
            />
          </div>
          <div style={{ fontSize: '12px', color: '#605e5c', lineHeight: '1.4' }}>
            <strong>Usage:</strong> Configure which SharePoint site to load data from. Leave empty to use the current site where this web part is deployed.
            This allows you to centralize your Adaptive Cards data in one location while deploying the Dashboard web part to multiple sites.
          </div>
        </div>

        {/* Context Information */}
        <div style={{ marginBottom: '20px', padding: '12px', background: '#f3f2f1', borderRadius: '6px', border: '1px solid #edebe9' }}>
          <div style={{ display: 'flex', alignItems: 'center', fontSize: '14px' }}>
            <span style={{ marginRight: '8px' }}>
              {this.isTeamsContext() ? '👥' : '📋'}
            </span>
            <span style={{ fontWeight: '600' }}>
              {this.isTeamsContext() ? 'Teams Context' : 'SharePoint Context'}
            </span>
            {this.state.customSiteUrl && (
              <span style={{ marginLeft: '12px', color: '#605e5c' }}>
                → {this.state.customSiteUrl}
              </span>
            )}
          </div>
        </div>

        {/* Filters */}
        {!loading && this.renderFilters()}

        {/* Charts */}
        {!loading && this.renderCharts()}

        {/* Loading State */}
        {loading && (
          <div className={styles.loading}>
            <div className={styles.spinner}></div>
            <span>Loading messages...</span>
          </div>
        )}

        {/* Error State */}
        {error && (
          <div className={styles.error}>
            <div className={styles.errorIcon}>⚠️</div>
            <div>
              <strong>Error loading dashboard:</strong>
              <br />
              {error}
              <div style={{ fontSize: '14px', marginTop: '8px' }}>
                <strong>Note:</strong> Dashboard is showing sample data below. To connect to live SharePoint data:
                <ol style={{ marginTop: '8px', paddingLeft: '20px' }}>
                  <li>Ensure the SharePoint lists exist (run setup-sharepoint-lists.ps1)</li>
                  <li>Verify permissions to access the configured site</li>
                  <li>Check that the site URL is correct</li>
                </ol>
              </div>
            </div>
          </div>
        )}

        {/* Stats Cards */}
        <div style={{ 
          display: 'grid', 
          gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', 
          gap: '16px', 
          marginBottom: '24px' 
        }}>
          <div style={{ 
            background: 'white', 
            padding: '20px', 
            borderRadius: '8px', 
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
            textAlign: 'center'
          }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>📬</div>
            <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#0078d4' }}>{filteredMessages.length}</div>
            <div style={{ fontSize: '12px', color: '#666' }}>Total Messages</div>
          </div>
          <div style={{ 
            background: 'white', 
            padding: '20px', 
            borderRadius: '8px', 
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
            textAlign: 'center'
          }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>✅</div>
            <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#107c10' }}>{filteredMessages.filter(m => this.isMessageRead(m)).length}</div>
            <div style={{ fontSize: '12px', color: '#666' }}>Read Messages</div>
          </div>
          <div style={{ 
            background: 'white', 
            padding: '20px', 
            borderRadius: '8px', 
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
            textAlign: 'center'
          }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>🔔</div>
            <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#d13438' }}>{filteredMessages.filter(m => !this.isMessageRead(m)).length}</div>
            <div style={{ fontSize: '12px', color: '#666' }}>Unread Messages</div>
          </div>
          <div style={{ 
            background: 'white', 
            padding: '20px', 
            borderRadius: '8px', 
            boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
            textAlign: 'center'
          }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>⚡</div>
            <div style={{ fontSize: '24px', fontWeight: 'bold', color: '#ff8c00' }}>{filteredMessages.filter(m => m.Priority === 'High').length}</div>
            <div style={{ fontSize: '12px', color: '#666' }}>High Priority</div>
          </div>
        </div>

        {/* Messages List */}
        <div className={styles.messagesContainer}>
          <div className={styles.messagesHeader}>
            <h2>📋 Your Messages</h2>
            <div className={styles.messageStats}>
              <span>Showing {filteredMessages.length} messages</span>
            </div>
          </div>
          
          {filteredMessages.length === 0 ? (
            <div className={styles.noMessages}>
              <div className={styles.noMessagesIcon}>📭</div>
              <h3>No messages found</h3>
              <p>{error ? 'Try adjusting your filters or check your data connection.' : 'You\'re all caught up!'}</p>
            </div>
          ) : (
            <div className={styles.messagesList}>
              {filteredMessages.map(message => this.renderMessage(message))}
            </div>
          )}
        </div>

        {/* Quick Action Buttons */}
        <div style={{
          marginTop: '30px',
          padding: '20px',
          background: 'white',
          borderRadius: '8px',
          boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
        }}>
          <h3 style={{ 
            marginBottom: '16px', 
            color: '#323130',
            fontSize: '16px',
            fontWeight: '600'
          }}>🚀 Quick Actions</h3>
          <div style={{
            display: 'flex',
            gap: '12px',
            flexWrap: 'wrap'
          }}>
            <button
              onClick={this.openTeamsMessageCreator}
              style={{
                padding: '12px 20px',
                background: '#0078d4',
                color: 'white',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                transition: 'all 0.2s ease'
              }}
              onMouseOver={(e) => {
                e.currentTarget.style.background = '#106ebe';
                e.currentTarget.style.transform = 'translateY(-1px)';
              }}
              onMouseOut={(e) => {
                e.currentTarget.style.background = '#0078d4';
                e.currentTarget.style.transform = 'translateY(0)';
              }}
            >
              📝 Create Teams Message
            </button>
            
            <button
              onClick={this.openManagerDashboard}
              style={{
                padding: '12px 20px',
                background: '#107c10',
                color: 'white',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                transition: 'all 0.2s ease'
              }}
              onMouseOver={(e) => {
                e.currentTarget.style.background = '#0e6e0e';
                e.currentTarget.style.transform = 'translateY(-1px)';
              }}
              onMouseOut={(e) => {
                e.currentTarget.style.background = '#107c10';
                e.currentTarget.style.transform = 'translateY(0)';
              }}
            >
              🎛️ Manager Dashboard
            </button>
            
            <button
              onClick={this.openMessageDiagnostics}
              style={{
                padding: '12px 20px',
                background: '#d13438',
                color: 'white',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer',
                fontSize: '14px',
                fontWeight: '600',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                transition: 'all 0.2s ease'
              }}
              onMouseOver={(e) => {
                e.currentTarget.style.background = '#b92b2b';
                e.currentTarget.style.transform = 'translateY(-1px)';
              }}
              onMouseOut={(e) => {
                e.currentTarget.style.background = '#d13438';
                e.currentTarget.style.transform = 'translateY(0)';
              }}
            >
              🔍 Message Diagnostics
            </button>
          </div>
        </div>
      </div>
    );
  }
}
