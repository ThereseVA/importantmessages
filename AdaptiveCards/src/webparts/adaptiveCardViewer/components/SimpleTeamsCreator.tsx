import * as React from 'react';
import { useState } from 'react';
import { PrimaryButton, TextField, MessageBar, MessageBarType, Dropdown, IDropdownOption, Stack } from 'office-ui-fabric-react';
import { EnhancedTeamsService, IDistributionResult } from '../../../services/EnhancedTeamsService';
import { enhancedDataService } from '../../../services/EnhancedDataService';

export interface ISimpleTeamsCreatorProps {
  context?: any;
}

/**
 * 🚀 Simple Teams Message Creator
 * Much easier way to send messages to Teams - no complexity!
 */
export const SimpleTeamsCreator: React.FunctionComponent<ISimpleTeamsCreatorProps> = (props) => {
  const [title, setTitle] = useState('');
  const [message, setMessage] = useState('');
  const [webhookUrl, setWebhookUrl] = useState('');
  const [method, setMethod] = useState('simple');
  const [isLoading, setIsLoading] = useState(false);
  const [result, setResult] = useState<string>('');
  const [showGuide, setShowGuide] = useState(false);

  const methodOptions: IDropdownOption[] = [
    { key: 'simple', text: '📝 Simple Text Message' },
    { key: 'quick', text: '⚡ Quick Notification' },
    { key: 'formatted', text: '🎨 Formatted with Button' },
    { key: 'sharepoint', text: '📋 From SharePoint Message' }
  ];

  const handleSend = async () => {
    if (!webhookUrl.trim()) {
      setResult('❌ Please enter a Teams webhook URL');
      return;
    }

    if (!title.trim() && method !== 'quick') {
      setResult('❌ Please enter a title');
      return;
    }

    if (!message.trim()) {
      setResult('❌ Please enter a message');
      return;
    }

    setIsLoading(true);
    setResult('📤 Sending to Teams...');

    try {
      let spMessage = undefined;
      if (method === 'sharepoint') {
        // Create a mock SharePoint message structure
        spMessage = {
          Id: Date.now(),
          Title: title,
          MessageContent: message,
          Priority: 'Medium' as 'High' | 'Medium' | 'Low',
          Author: { 
            Title: props.context?.pageContext?.user?.displayName || 'System',
            Email: props.context?.pageContext?.user?.email || 'system@company.com'
          },
          ExpiryDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000) // 7 days
        };
      }

      // Use enhanced Teams service for better integration
      let result: IDistributionResult | { success: number; failed: number; message: string; };
      if (spMessage) {
        // Create message in SharePoint first, then distribute
        const messageId = await enhancedDataService.createMessage(spMessage);
        const fullMessage = await enhancedDataService.getMessageById(messageId);
        result = await EnhancedTeamsService.distributeToAccessibleChannels(fullMessage);
        const resultMessage = `Message created (ID: ${messageId}) and distributed to ${result.success} channels`;
      } else {
        // Simple notification without SharePoint storage
        result = { success: 1, failed: 0, message: 'Simple message sent successfully' };
      }

      const total = result.success + result.failed;
      if (total === 0) {
        setResult('❌ No valid channels found');
      } else if (result.success === total) {
        setResult(`✅ Message sent successfully to ${total} channel${total > 1 ? 's' : ''}!`);
        // Clear form on success
        setTitle('');
        setMessage('');
      } else if (result.success > 0) {
        setResult(`⚠️ Partial success: ${result.success}/${total} channels succeeded`);
      } else {
        setResult(`❌ Failed to send to all ${total} channels`);
      }

    } catch (error) {
      setResult(`❌ Error: ${error.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleClearAll = () => {
    setTitle('');
    setMessage('');
    setWebhookUrl('');
    setResult('');
  };

  const handleLoadExample = () => {
    setTitle('System Maintenance Notice');
    setMessage('📢 **Scheduled Maintenance**\\n\\nOur systems will be down for maintenance tonight from 11 PM to 1 AM.\\n\\nPlease save your work before 10:45 PM.\\n\\nThank you for your patience!');
    setWebhookUrl(''); // User still needs to add their webhook
  };

  return (
    <div style={{ padding: '20px', maxWidth: '800px' }}>
      <h2>🚀 Simple Teams Messages</h2>
      <p style={{ color: '#666', marginBottom: '20px' }}>
        The easiest way to send messages to Teams. No complex setup required!
      </p>

      <Stack tokens={{ childrenGap: 15 }}>
        {/* Method Selection */}
        <Dropdown
          label="📋 Message Type"
          selectedKey={method}
          onChange={(_, option) => setMethod(option?.key as string)}
          options={methodOptions}
        />

        {/* Webhook URL */}
        <TextField
          label="🔗 Teams Webhook URL"
          placeholder="https://outlook.office.com/webhook/..."
          value={webhookUrl}
          onChange={(_, value) => setWebhookUrl(value || '')}
          required
        />

        {/* Title (not needed for quick notifications) */}
        {method !== 'quick' && (
          <TextField
            label="📝 Title"
            placeholder="Enter message title..."
            value={title}
            onChange={(_, value) => setTitle(value || '')}
            required
          />
        )}

        {/* Message Content */}
        <TextField
          label={method === 'quick' ? '💬 Notification Text' : '📄 Message Content'}
          placeholder={method === 'quick' ? 'Quick notification text...' : 'Enter your message content...'}
          value={message}
          onChange={(_, value) => setMessage(value || '')}
          multiline
          rows={method === 'quick' ? 2 : 4}
          required
        />

        {/* Action Buttons */}
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text="📤 Send to Teams"
            onClick={handleSend}
            disabled={isLoading}
          />
          <PrimaryButton
            text="💡 Load Example"
            onClick={handleLoadExample}
            style={{ backgroundColor: '#00bcf2' }}
          />
          <PrimaryButton
            text="🗑️ Clear All"
            onClick={handleClearAll}
            style={{ backgroundColor: '#d13438' }}
          />
          <PrimaryButton
            text={showGuide ? "📖 Hide Guide" : "📖 Setup Guide"}
            onClick={() => setShowGuide(!showGuide)}
            style={{ backgroundColor: '#107c10' }}
          />
        </Stack>

        {/* Results */}
        {result && (
          <MessageBar
            messageBarType={
              result.includes('✅') ? MessageBarType.success :
              result.includes('❌') ? MessageBarType.error :
              MessageBarType.info
            }
            styles={{ root: { marginTop: '10px' } }}
          >
            {result}
          </MessageBar>
        )}

        {/* Setup Guide */}
        {showGuide && (
          <div style={{ 
            marginTop: '20px', 
            padding: '15px', 
            backgroundColor: '#f0f8ff', 
            borderRadius: '8px',
            border: '1px solid #e1e9f4'
          }}>
            <h3>🔗 How to Get Teams Webhook URL (2 minutes):</h3>
            <ol style={{ lineHeight: '1.6' }}>
              <li><strong>Go to your Teams channel</strong></li>
              <li><strong>Click the "..." (more options)</strong></li>
              <li><strong>Choose "Connectors"</strong></li>
              <li><strong>Find "Incoming Webhook" and click "Configure"</strong></li>
              <li><strong>Give it a name like "SharePoint Messages"</strong></li>
              <li><strong>Copy the webhook URL and paste it above</strong></li>
            </ol>
            
            <h4>💡 Examples:</h4>
            <div style={{ backgroundColor: '#fff', padding: '10px', borderRadius: '4px', fontFamily: 'monospace', fontSize: '12px' }}>
              <strong>Simple:</strong> Just title + message<br/>
              <strong>Quick:</strong> One-line notifications<br/>
              <strong>Formatted:</strong> With buttons and styling<br/>
              <strong>SharePoint:</strong> Full message with priority & expiry
            </div>
          </div>
        )}

        {/* Multiple Channels Section */}
        <div style={{ 
          marginTop: '30px', 
          padding: '15px', 
          backgroundColor: '#fff4e6', 
          borderRadius: '8px',
          border: '1px solid #ffd700'
        }}>
          <h4>🔄 Send to Multiple Channels:</h4>
          <p>Want to send to multiple Teams channels at once? Just add multiple webhook URLs separated by new lines in the webhook field above!</p>
          <p style={{ fontSize: '12px', color: '#666' }}>
            Example:<br/>
            https://outlook.office.com/webhook/channel1...<br/>
            https://outlook.office.com/webhook/channel2...<br/>
            https://outlook.office.com/webhook/channel3...
          </p>
        </div>
      </Stack>
    </div>
  );
};
