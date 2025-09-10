import * as React from 'react';
import { useState, useEffect } from 'react';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { MessageBar, MessageBarType, PrimaryButton, DefaultButton } from 'office-ui-fabric-react';

export interface IMarkReadPageProps {
  messageId: number;
  source?: string;
  // dataService no longer needed - using global enhancedDataService
}

export const MarkReadPage: React.FunctionComponent<IMarkReadPageProps> = (props) => {
  const [status, setStatus] = useState<'loading' | 'success' | 'error' | 'already-read'>('loading');
  const [message, setMessage] = useState<any>(null);
  const [error, setError] = useState<string>('');

  useEffect(() => {
    markAsRead();
  }, []);

  const markAsRead = async () => {
    try {
      setStatus('loading');
      
      if (!enhancedDataService.getCurrentUser()) {
        setError('Data service not available');
        setStatus('error');
        return;
      }

      // First, get the message details
      const messageData = await enhancedDataService.getMessageById(props.messageId);
      setMessage(messageData);

      // Check if already read
      const alreadyRead = await enhancedDataService.hasUserReadMessage(props.messageId);
      if (alreadyRead) {
        setStatus('already-read');
        return;
      }

      // Mark as read
      await enhancedDataService.markMessageAsRead(props.messageId);
      setStatus('success');

    } catch (err) {
      console.error('Error marking message as read:', err);
      setError(err.message || 'Failed to mark message as read');
      setStatus('error');
    }
  };

  const goToDashboard = () => {
    // Always use the correct subsite URL regardless of current context
    window.location.href = `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/SitePages/Dashboard.aspx`;
  };

  const goToMessages = () => {
    // Always use the correct subsite URL regardless of current context
    window.location.href = `https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/Lists/Important%20Messages/AllItems.aspx`;
  };

  const goToTeams = () => {
    // If coming from Teams, try to go back
    if (props.source === 'teams') {
      window.close(); // This might work in Teams context
    }
  };

  return (
    <div style={{ padding: '20px', maxWidth: '600px', margin: '0 auto' }}>
      <h2>📨 Message Action</h2>

      {message && (
        <div style={{ 
          background: '#f8f9fa', 
          padding: '15px', 
          borderRadius: '8px', 
          marginBottom: '20px',
          border: '1px solid #dee2e6'
        }}>
          <h3>{message.Title}</h3>
          <p>{message.MessageContent}</p>
          <small>From: {message.Author?.Title} | Priority: {message.Priority}</small>
        </div>
      )}

      {status === 'loading' && (
        <MessageBar messageBarType={MessageBarType.info}>
          📤 Processing your request...
        </MessageBar>
      )}

      {status === 'success' && (
        <div>
          <MessageBar messageBarType={MessageBarType.success}>
            ✅ Message marked as read successfully! Thank you for confirming.
          </MessageBar>
          <div style={{ marginTop: '20px' }}>
            <PrimaryButton text="📊 View Dashboard" onClick={goToDashboard} style={{ marginRight: '10px' }} />
            <DefaultButton text="📋 All Messages" onClick={goToMessages} style={{ marginRight: '10px' }} />
            {props.source === 'teams' && (
              <DefaultButton text="↩️ Back to Teams" onClick={goToTeams} />
            )}
          </div>
        </div>
      )}

      {status === 'already-read' && (
        <div>
          <MessageBar messageBarType={MessageBarType.warning}>
            ℹ️ You have already marked this message as read.
          </MessageBar>
          <div style={{ marginTop: '20px' }}>
            <PrimaryButton text="📊 View Dashboard" onClick={goToDashboard} style={{ marginRight: '10px' }} />
            <DefaultButton text="📋 All Messages" onClick={goToMessages} style={{ marginRight: '10px' }} />
            {props.source === 'teams' && (
              <DefaultButton text="↩️ Back to Teams" onClick={goToTeams} />
            )}
          </div>
        </div>
      )}

      {status === 'error' && (
        <div>
          <MessageBar messageBarType={MessageBarType.error}>
            ❌ Error: {error}
          </MessageBar>
          <div style={{ marginTop: '20px' }}>
            <PrimaryButton text="🔄 Try Again" onClick={markAsRead} style={{ marginRight: '10px' }} />
            <DefaultButton text="📊 View Dashboard" onClick={goToDashboard} style={{ marginRight: '10px' }} />
            <DefaultButton text="📋 All Messages" onClick={goToMessages} />
          </div>
        </div>
      )}

      <div style={{ 
        marginTop: '30px', 
        padding: '15px', 
        background: '#e8f4fd', 
        borderRadius: '8px',
        fontSize: '14px'
      }}>
        <h4>📚 About Message Tracking:</h4>
        <ul>
          <li>✅ Your read confirmation is logged in SharePoint</li>
          <li>📊 Administrators can view read statistics on the dashboard</li>
          <li>🔒 Only you and authorized personnel can see your read status</li>
          <li>📱 This works from Teams, email, or SharePoint</li>
        </ul>
      </div>
    </div>
  );
};
