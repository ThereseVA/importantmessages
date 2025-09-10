import * as React from 'react';
import { useState, useEffect } from 'react';
import { 
  DetailsList, 
  IColumn, 
  SelectionMode, 
  PrimaryButton,
  DefaultButton,
  SearchBox,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { enhancedDataService, IMessage } from '../../../services/EnhancedDataService';

export interface IEmployeeMessageListProps {
  // Props simplified since we use global enhancedDataService
}

export const EmployeeMessageList: React.FunctionComponent<IEmployeeMessageListProps> = (props) => {
  const [messages, setMessages] = useState<IMessage[]>([]);
  const [filteredMessages, setFilteredMessages] = useState<IMessage[]>([]);
  const [loading, setLoading] = useState(true);
  const [selectedMessage, setSelectedMessage] = useState<IMessage | null>(null);
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [searchText, setSearchText] = useState('');
  const [priorityFilter, setPriorityFilter] = useState('All');
  const [result, setResult] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);

  const priorityOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Priorities' },
    { key: 'High', text: '🚨 High Priority' },
    { key: 'Medium', text: '⚠️ Medium Priority' },
    { key: 'Low', text: 'ℹ️ Low Priority' }
  ];

  useEffect(() => {
    loadMessages();
  }, []);

  useEffect(() => {
    applyFilters();
  }, [messages, searchText, priorityFilter]);

  const loadMessages = async () => {
    setLoading(true);
    try {
      // Get messages filtered for current user (employee view)
      const userMessages = await enhancedDataService.getMessagesForCurrentUser();
      setMessages(userMessages);
      setResult(null);
    } catch (error) {
      console.error('Error loading messages:', error);
      setResult({ 
        type: 'error', 
        message: `❌ Failed to load messages: ${error.message || 'Unknown error'}` 
      });
    } finally {
      setLoading(false);
    }
  };

  const applyFilters = () => {
    let filtered = [...messages];

    // Apply search filter
    if (searchText.trim()) {
      const searchLower = searchText.toLowerCase();
      filtered = filtered.filter(msg => 
        msg.Title.toLowerCase().includes(searchLower) ||
        msg.MessageContent.toLowerCase().includes(searchLower)
      );
    }

    // Apply priority filter
    if (priorityFilter !== 'All') {
      filtered = filtered.filter(msg => msg.Priority === priorityFilter);
    }

    // Sort by priority and creation date
    filtered.sort((a, b) => {
      const priorityOrder = { 'High': 3, 'Medium': 2, 'Low': 1 };
      const priorityDiff = priorityOrder[b.Priority] - priorityOrder[a.Priority];
      if (priorityDiff !== 0) return priorityDiff;
      
      return new Date(b.Created).getTime() - new Date(a.Created).getTime();
    });

    setFilteredMessages(filtered);
  };

  const handleMarkAsRead = async (message: IMessage) => {
    try {
      await enhancedDataService.markMessageAsRead(message.Id);
      setResult({ 
        type: 'success', 
        message: `✅ Message "${message.Title}" marked as read` 
      });
      
      // Refresh messages to update read status
      loadMessages();
    } catch (error) {
      setResult({ 
        type: 'error', 
        message: `❌ Failed to mark as read: ${error.message}` 
      });
    }
  };

  const handleViewMessage = (message: IMessage) => {
    setSelectedMessage(message);
    setIsPanelOpen(true);
  };

  const checkIfRead = async (messageId: number): Promise<boolean> => {
    try {
      return await enhancedDataService.hasUserReadMessage(messageId);
    } catch (error) {
      return false;
    }
  };

  const columns: IColumn[] = [
    {
      key: 'priority',
      name: 'Priority',
      fieldName: 'Priority',
      minWidth: 80,
      maxWidth: 80,
      onRender: (item: IMessage) => {
        const priorityIcons = {
          'High': '🚨',
          'Medium': '⚠️',
          'Low': 'ℹ️'
        };
        return `${priorityIcons[item.Priority]} ${item.Priority}`;
      }
    },
    {
      key: 'title',
      name: 'Message Title',
      fieldName: 'Title',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: IMessage) => (
        <div>
          <strong>{item.Title}</strong>
          <div style={{ fontSize: '12px', color: '#666' }}>
            {new Date(item.Created).toLocaleDateString('sv-SE')} {new Date(item.Created).toLocaleTimeString('sv-SE')}
          </div>
        </div>
      )
    },
    {
      key: 'content',
      name: 'Content Preview',
      fieldName: 'MessageContent',
      minWidth: 250,
      isResizable: true,
      onRender: (item: IMessage) => {
        // Strip HTML and truncate
        const textContent = item.MessageContent.replace(/<[^>]*>/g, '');
        const preview = textContent.length > 100 ? textContent.substring(0, 100) + '...' : textContent;
        return <span>{preview}</span>;
      }
    },
    {
      key: 'targetAudience',
      name: 'Target Audience',
      fieldName: 'TargetAudience',
      minWidth: 120,
      maxWidth: 150
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 120,
      onRender: (item: IMessage) => (
        <div style={{ display: 'flex', gap: '8px' }}>
          <DefaultButton 
            text="View" 
            onClick={() => handleViewMessage(item)}
            styles={{ root: { minWidth: '50px' } }}
          />
          <PrimaryButton 
            text="Mark Read" 
            onClick={() => handleMarkAsRead(item)}
            styles={{ root: { minWidth: '70px' } }}
          />
        </div>
      )
    }
  ];

  const renderMessagePanel = () => (
    <Panel
      isOpen={isPanelOpen}
      onDismiss={() => setIsPanelOpen(false)}
      type={PanelType.medium}
      headerText={selectedMessage?.Title || 'Message Details'}
      closeButtonAriaLabel="Close"
    >
      {selectedMessage && (
        <div style={{ padding: '16px' }}>
          <div style={{ marginBottom: '16px' }}>
            <strong>Priority:</strong> {selectedMessage.Priority === 'High' ? '🚨' : selectedMessage.Priority === 'Medium' ? '⚠️' : 'ℹ️'} {selectedMessage.Priority}
          </div>
          
          <div style={{ marginBottom: '16px' }}>
            <strong>Target Audience:</strong> {selectedMessage.TargetAudience}
          </div>
          
          <div style={{ marginBottom: '16px' }}>
            <strong>Created:</strong> {new Date(selectedMessage.Created).toLocaleString('sv-SE')}
          </div>
          
          <div style={{ marginBottom: '24px' }}>
            <strong>Message Content:</strong>
            <div 
              style={{ 
                marginTop: '8px', 
                padding: '12px', 
                border: '1px solid #ddd', 
                borderRadius: '4px',
                backgroundColor: '#f9f9f9' 
              }}
              dangerouslySetInnerHTML={{ __html: selectedMessage.MessageContent }}
            />
          </div>
          
          <PrimaryButton 
            text="Mark as Read" 
            onClick={() => {
              handleMarkAsRead(selectedMessage);
              setIsPanelOpen(false);
            }}
          />
        </div>
      )}
    </Panel>
  );

  if (loading) {
    return (
      <div style={{ textAlign: 'center', padding: '40px' }}>
        <Spinner size={SpinnerSize.large} label="Loading your messages..." />
      </div>
    );
  }

  return (
    <div style={{ padding: '20px' }}>
      <h2>📨 Your Important Messages</h2>
      <p>Messages targeted to you and your role</p>

      {result && (
        <MessageBar
          messageBarType={result.type === 'success' ? MessageBarType.success : 
                        result.type === 'error' ? MessageBarType.error : MessageBarType.info}
          styles={{ root: { marginBottom: '16px' } }}
          onDismiss={() => setResult(null)}
        >
          {result.message}
        </MessageBar>
      )}

      {/* Filters */}
      <div style={{ display: 'flex', gap: '16px', marginBottom: '20px', alignItems: 'flex-end' }}>
        <SearchBox
          placeholder="Search messages..."
          value={searchText}
          onChange={(_, newValue) => setSearchText(newValue || '')}
          styles={{ root: { width: '300px' } }}
        />
        
        <Dropdown
          placeholder="Filter by priority"
          selectedKey={priorityFilter}
          options={priorityOptions}
          onChange={(_, option) => setPriorityFilter(option?.key as string || 'All')}
          styles={{ dropdown: { width: '150px' } }}
        />
        
        <DefaultButton 
          text="Refresh" 
          iconProps={{ iconName: 'Refresh' }}
          onClick={loadMessages} 
        />
      </div>

      {/* Messages List */}
      {filteredMessages.length === 0 ? (
        <MessageBar messageBarType={MessageBarType.info}>
          {messages.length === 0 
            ? "📭 No messages found. You're all caught up!" 
            : "📭 No messages match your current filters."
          }
        </MessageBar>
      ) : (
        <DetailsList
          items={filteredMessages}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={0}
          styles={{
            root: { border: '1px solid #ddd' }
          }}
        />
      )}

      {/* Message Detail Panel */}
      {renderMessagePanel()}

      {/* Help Section */}
      <div style={{ marginTop: '30px', padding: '15px', backgroundColor: '#e8f4fd', borderRadius: '8px' }}>
        <h4>💡 Employee Message Center</h4>
        <ul>
          <li><strong>📋 View Messages:</strong> Messages targeted to your role and groups</li>
          <li><strong>🔍 Search & Filter:</strong> Find specific messages quickly</li>
          <li><strong>✅ Mark as Read:</strong> Confirm you've seen important information</li>
          <li><strong>📊 Priority Levels:</strong> High 🚨, Medium ⚠️, Low ℹ️</li>
        </ul>
      </div>
    </div>
  );
};
