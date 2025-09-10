import * as React from 'react';
import { useState, useEffect } from 'react';
import { enhancedDataService, IMessage } from '../../../services/EnhancedDataService';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  IColumn, 
  SelectionMode,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  DefaultButton,
  SearchBox,
  Dropdown,
  IDropdownOption,
  ProgressIndicator,
  Panel,
  PanelType
} from 'office-ui-fabric-react';

export interface IManagerDashboardProps {
  // Props are simplified since we use the global enhancedDataService
}

interface IMessageWithStats extends IMessage {
  totalReads: number;
  uniqueReaders: number;
  readPercentage: number;
  lastReadDate?: Date;
  readStatus: 'Not Started' | 'In Progress' | 'Completed';
  notReadUsers: string[];
}

export const ManagerDashboard: React.FunctionComponent<IManagerDashboardProps> = (props) => {
  const [messages, setMessages] = useState<IMessageWithStats[]>([]);
  const [filteredMessages, setFilteredMessages] = useState<IMessageWithStats[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedMessage, setSelectedMessage] = useState<IMessageWithStats | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState('All');
  const [sourceFilter, setSourceFilter] = useState('All');

  console.log('🔧 ManagerDashboard: Component rendered, enhanced data service available:', !!enhancedDataService);

  const statusOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Messages' },
    { key: 'Not Started', text: '🔴 Not Started (0% read)' },
    { key: 'In Progress', text: '🟡 In Progress (1-99% read)' },
    { key: 'Completed', text: '🟢 Completed (100% read)' }
  ];

  const sourceOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Sources' },
    { key: 'Teams', text: '👥 Teams Messages' },
    { key: 'SharePoint', text: '📋 SharePoint Messages' },
    { key: 'Outlook', text: '📧 Outlook Emails' }
  ];

  useEffect(() => {
    if (!enhancedDataService.getCurrentUser()) {
      console.error('🔧 ManagerDashboard: Enhanced data service not initialized');
      setError('Enhanced data service not initialized');
      setLoading(false);
      return;
    }
    
    loadMessages();
  }, []);

  useEffect(() => {
    applyFilters();
  }, [messages, searchText, statusFilter, sourceFilter]);

  const loadMessages = async () => {
    setLoading(true);
    setError(null);
    try {
      console.log('🔧 ManagerDashboard: Loading all messages for manager view');
      
      if (!enhancedDataService.getCurrentUser()) {
        throw new Error('Enhanced data service not available');
      }
      
      // Managers should see ALL messages, not filtered ones
      const allMessages = await enhancedDataService.getActiveMessages();
      console.log('🔧 ManagerDashboard: Retrieved', allMessages.length, 'messages');
      
      const messagesWithStats = await Promise.all(
        allMessages.map(async (message) => {
          try {
            // TODO: Implement read stats in EnhancedDataService
            const stats = { 
              totalReads: 0, 
              uniqueReaders: 0, 
              readPercentage: 0, 
              unreadCount: 0,
              readActions: [] as any[]
            };
            // const stats = await enhancedDataService.getMessageReadStats(message.Id);
            
            // Calculate read percentage (you'll need to know total target users)
            // For now, we'll estimate based on "All Users" = 100, specific groups = smaller numbers
            const estimatedTargetUsers = message.TargetAudience === 'All Users' ? 100 : 
                                       message.TargetAudience.includes('Department') ? 25 : 50;
            
            const readPercentage = Math.min((stats.uniqueReaders / estimatedTargetUsers) * 100, 100);
            
            let readStatus: 'Not Started' | 'In Progress' | 'Completed';
            if (readPercentage === 0) readStatus = 'Not Started';
            else if (readPercentage < 100) readStatus = 'In Progress';
            else readStatus = 'Completed';

            return {
              ...message,
              totalReads: stats.totalReads,
              uniqueReaders: stats.uniqueReaders,
              readPercentage: Math.round(readPercentage),
              lastReadDate: stats.readActions.length > 0 ? stats.readActions[0].ReadTimestamp : undefined,
              readStatus,
              notReadUsers: [] // You could calculate this based on your user directory
            } as IMessageWithStats;
          } catch (error) {
            console.error('🔧 ManagerDashboard: Error getting stats for message', message.Id, error);
            // Return message with default stats if stats retrieval fails
            return {
              ...message,
              totalReads: 0,
              uniqueReaders: 0,
              readPercentage: 0,
              readStatus: 'Not Started' as const,
              notReadUsers: []
            } as IMessageWithStats;
          }
        })
      );

      console.log('🔧 ManagerDashboard: Processed', messagesWithStats.length, 'messages with stats');
      setMessages(messagesWithStats);
    } catch (error) {
      console.error('🔧 ManagerDashboard: Error loading messages with stats:', error);
      setError(`Failed to load messages: ${error.message || 'Unknown error'}`);
      // Set empty array instead of leaving in loading state
      setMessages([]);
    } finally {
      setLoading(false);
    }
  };

  const applyFilters = () => {
    let filtered = [...messages];

    // Apply search filter
    if (searchText) {
      filtered = filtered.filter(msg => 
        msg.Title.toLowerCase().includes(searchText.toLowerCase()) ||
        (msg.MessageContent && msg.MessageContent.toLowerCase().includes(searchText.toLowerCase()))
      );
    }

    // Apply status filter
    if (statusFilter !== 'All') {
      filtered = filtered.filter(msg => msg.readStatus === statusFilter);
    }

    // Apply source filter
    if (sourceFilter !== 'All') {
      filtered = filtered.filter(msg => (msg.Source || 'SharePoint') === sourceFilter);
    }

    setFilteredMessages(filtered);
  };

  const columns: IColumn[] = [
    {
      key: 'status',
      name: 'Status',
      fieldName: 'readStatus',
      minWidth: 80,
      maxWidth: 120,
      onRender: (item: IMessageWithStats) => {
        const icon = item.readStatus === 'Completed' ? '🟢' : 
                    item.readStatus === 'In Progress' ? '🟡' : '🔴';
        return <span>{icon} {item.readStatus}</span>;
      }
    },
    {
      key: 'source',
      name: 'Source',
      fieldName: 'Source',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: IMessageWithStats) => {
        const source = item.Source || 'SharePoint';
        const icon = source === 'Teams' ? '👥' : source === 'Outlook' ? '📧' : '📋';
        return <span>{icon} {source}</span>;
      }
    },
    {
      key: 'priority',
      name: 'Priority',
      fieldName: 'Priority',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: IMessageWithStats) => {
        const icon = item.Priority === 'High' ? '🚨' : 
                    item.Priority === 'Medium' ? '⚠️' : 'ℹ️';
        return <span>{icon} {item.Priority}</span>;
      }
    },
    {
      key: 'title',
      name: 'Message Title',
      fieldName: 'Title',
      minWidth: 200,
      maxWidth: 400,
      isResizable: true,
      onRender: (item: IMessageWithStats) => (
        <div>
          <strong>{item.Title}</strong>
          <div style={{ fontSize: '12px', color: '#666' }}>
            {item.MessageContent ? item.MessageContent.substring(0, 100) + '...' : 'No content'}
          </div>
        </div>
      )
    },
    {
      key: 'readProgress',
      name: 'Read Progress',
      minWidth: 150,
      maxWidth: 200,
      onRender: (item: IMessageWithStats) => (
        <div>
          <ProgressIndicator 
            percentComplete={item.readPercentage / 100} 
            description={`${item.readPercentage}% (${item.uniqueReaders} users)`}
          />
        </div>
      )
    },
    {
      key: 'created',
      name: 'Created',
      fieldName: 'Created',
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: IMessageWithStats) => item.Created.toLocaleDateString()
    },
    {
      key: 'target',
      name: 'Target Audience',
      fieldName: 'TargetAudience',
      minWidth: 120,
      maxWidth: 180
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: IMessageWithStats) => (
        <DefaultButton 
          text="👥 View Details" 
          onClick={() => {
            setSelectedMessage(item);
            setIsPanelOpen(true);
          }}
        />
      )
    }
  ];

  const getOverallStats = () => {
    const total = messages.length;
    const completed = messages.filter(m => m.readStatus === 'Completed').length;
    const inProgress = messages.filter(m => m.readStatus === 'In Progress').length;
    const notStarted = messages.filter(m => m.readStatus === 'Not Started').length;
    const avgReadRate = total > 0 ? Math.round(messages.reduce((sum, m) => sum + m.readPercentage, 0) / total) : 0;

    return { total, completed, inProgress, notStarted, avgReadRate };
  };

  const stats = getOverallStats();

  if (error) {
    return (
      <div style={{ padding: '20px' }}>
        <h2>📊 Manager Dashboard - Error</h2>
        <MessageBar messageBarType={MessageBarType.error}>
          {error}
        </MessageBar>
        <div style={{ marginTop: '20px' }}>
          <DefaultButton 
            text="Retry" 
            onClick={loadMessages}
            iconProps={{ iconName: 'Refresh' }}
          />
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div style={{ padding: '20px' }}>
        <h2>📊 Manager Dashboard - Loading...</h2>
        <ProgressIndicator description="Loading message statistics..." />
      </div>
    );
  }

  return (
    <div style={{ padding: '20px' }}>
      <div style={{ marginBottom: '20px' }}>
        <h2 style={{ color: '#323130', marginBottom: '10px' }}>📊 Unified Message Dashboard</h2>
        <div style={{ color: '#605e5c', fontSize: '14px' }}>
          View all messages from Teams and SharePoint with read tracking analytics
        </div>
      </div>
      
      {/* Overall Statistics */}
      <div style={{ 
        display: 'grid', 
        gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', 
        gap: '15px', 
        marginBottom: '20px' 
      }}>
        <div style={{ background: '#f0f8ff', padding: '15px', borderRadius: '8px', textAlign: 'center' }}>
          <h3 style={{ margin: '0 0 5px 0', color: '#0078d4' }}>{stats.total}</h3>
          <div>Total Messages</div>
        </div>
        <div style={{ background: '#f0fff0', padding: '15px', borderRadius: '8px', textAlign: 'center' }}>
          <h3 style={{ margin: '0 0 5px 0', color: '#107c10' }}>🟢 {stats.completed}</h3>
          <div>Fully Read</div>
        </div>
        <div style={{ background: '#fffbf0', padding: '15px', borderRadius: '8px', textAlign: 'center' }}>
          <h3 style={{ margin: '0 0 5px 0', color: '#f7630c' }}>🟡 {stats.inProgress}</h3>
          <div>Partially Read</div>
        </div>
        <div style={{ background: '#fff0f0', padding: '15px', borderRadius: '8px', textAlign: 'center' }}>
          <h3 style={{ margin: '0 0 5px 0', color: '#d13438' }}>🔴 {stats.notStarted}</h3>
          <div>Not Started</div>
        </div>
        <div style={{ background: '#f8f9fa', padding: '15px', borderRadius: '8px', textAlign: 'center' }}>
          <h3 style={{ margin: '0 0 5px 0', color: '#323130' }}>{stats.avgReadRate}%</h3>
          <div>Avg Read Rate</div>
        </div>
      </div>

      {/* Filters */}
      <div style={{ display: 'flex', gap: '15px', marginBottom: '20px', alignItems: 'end' }}>
        <SearchBox
          placeholder="Search messages..."
          value={searchText}
          onChange={(_, newValue) => setSearchText(newValue || '')}
          styles={{ root: { width: '300px' } }}
        />
        <Dropdown
          label="Filter by Status"
          selectedKey={statusFilter}
          onChange={(_, option) => setStatusFilter(option?.key as string || 'All')}
          options={statusOptions}
          styles={{ root: { width: '200px' } }}
        />
        <Dropdown
          label="Filter by Source"
          selectedKey={sourceFilter}
          onChange={(_, option) => setSourceFilter(option?.key as string || 'All')}
          options={sourceOptions}
          styles={{ root: { width: '180px' } }}
        />
        <PrimaryButton text="🔄 Refresh" onClick={loadMessages} />
      </div>

      {/* Messages List */}
      <DetailsList
        items={filteredMessages}
        columns={columns}
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
        isHeaderVisible={true}
      />

      {/* Details Panel */}
      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        type={PanelType.medium}
        headerText={selectedMessage ? `📊 Details: ${selectedMessage.Title}` : ''}
      >
        {selectedMessage && (
          <div style={{ padding: '10px' }}>
            <MessageBar messageBarType={
              selectedMessage.readStatus === 'Completed' ? MessageBarType.success :
              selectedMessage.readStatus === 'In Progress' ? MessageBarType.warning :
              MessageBarType.error
            }>
              Status: {selectedMessage.readStatus} - {selectedMessage.readPercentage}% read by users
            </MessageBar>

            <div style={{ marginTop: '20px' }}>
              <h3>📋 Message Details</h3>
              <p><strong>Content:</strong> {selectedMessage.MessageContent || 'No content'}</p>
              <p><strong>Priority:</strong> {selectedMessage.Priority}</p>
              <p><strong>Target Audience:</strong> {selectedMessage.TargetAudience}</p>
              <p><strong>Created:</strong> {selectedMessage.Created.toLocaleString()}</p>
              <p><strong>Expires:</strong> {selectedMessage.ExpiryDate.toLocaleString()}</p>
            </div>

            <div style={{ marginTop: '20px' }}>
              <h3>📊 Read Statistics</h3>
              <p><strong>Total Reads:</strong> {selectedMessage.totalReads}</p>
              <p><strong>Unique Readers:</strong> {selectedMessage.uniqueReaders}</p>
              <p><strong>Read Percentage:</strong> {selectedMessage.readPercentage}%</p>
              {selectedMessage.lastReadDate && (
                <p><strong>Last Read:</strong> {selectedMessage.lastReadDate.toLocaleString()}</p>
              )}
            </div>

            <div style={{ marginTop: '20px' }}>
              <h3>🎯 Actions</h3>
              <div style={{ display: 'flex', gap: '10px', flexDirection: 'column' }}>
                <PrimaryButton 
                  text="📤 Resend to Teams" 
                  onClick={() => {
                    // Implement resend functionality
                    alert('Resending message to Teams channels...');
                  }}
                />
                <DefaultButton 
                  text="📧 Send Reminder Email" 
                  onClick={() => {
                    // Implement email reminder
                    alert('Sending reminder emails to non-readers...');
                  }}
                />
                <DefaultButton 
                  text="📋 Export Read Report" 
                  onClick={() => {
                    // Implement export functionality
                    alert('Exporting detailed read report...');
                  }}
                />
              </div>
            </div>
          </div>
        )}
      </Panel>
    </div>
  );
};
