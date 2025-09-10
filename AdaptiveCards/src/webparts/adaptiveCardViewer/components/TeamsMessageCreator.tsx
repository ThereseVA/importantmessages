import * as React from 'react';
import { useState, useEffect } from 'react';
import { PrimaryButton, DefaultButton, TextField, Dropdown, MessageBar, MessageBarType, IDropdownOption, Label, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { EnhancedTeamsService } from '../../../services/EnhancedTeamsService';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { SiteSelector } from './SiteSelector';

export interface ITeamsMessageCreatorProps {
  context?: any;
  dataService?: any;
  onMessageCreated?: (messageId: number) => void;
}

export const TeamsMessageCreator: React.FunctionComponent<ITeamsMessageCreatorProps> = (props) => {
  console.log('🎯 TeamsMessageCreator component started');
  console.log('🎯 Props received:', props);
  console.log('🎯 Context available:', !!props.context);
  
  // Manager permission state
  const [isManager, setIsManager] = useState<boolean | null>(null);
  const [isCheckingPermissions, setIsCheckingPermissions] = useState(true);
  
  // Initialize with current SharePoint context if available
  const [currentSite, setCurrentSite] = useState<string>(
    props.context?.pageContext?.web?.absoluteUrl || 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken'
  );
  const [currentSiteName, setCurrentSiteName] = useState<string>(
    props.context?.pageContext?.web?.title || 'Current Site'
  );
  const [formData, setFormData] = useState({
    title: '',
    content: '',
    priority: 'Medium',
    targetAudience: 'Teams Channel',
    expiryDays: '7',
    distributionChannels: [] as string[],
    useEmailIntegration: false // New option for email-based Teams integration
  });
  
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [result, setResult] = useState<{ type: 'success' | 'error' | 'info'; message: string } | null>(null);
  const [webhookUrls, setWebhookUrls] = useState<string>('');

  console.log('🎯 TeamsMessageCreator state initialized');
  console.log('🎯 Form data:', formData);

  // Check manager permissions on component mount
  useEffect(() => {
    const checkManagerPermissions = async () => {
      if (!props.context) {
        console.warn('🎯 No context available for permission check');
        setIsCheckingPermissions(false);
        setIsManager(false);
        return;
      }

      try {
        console.log('🎯 Checking manager permissions...');
        
        // Initialize the enhanced data service if not already done
        try {
          await enhancedDataService.initialize(props.context);
        } catch (initError) {
          console.warn('🎯 Service already initialized or initialization failed:', initError);
        }
        
        // Check if current user is a manager using the Managers SharePoint list
        const managerStatus = await enhancedDataService.isCurrentUserManager();
        console.log('🎯 Manager status:', managerStatus);
        
        setIsManager(managerStatus);
        setIsCheckingPermissions(false);
      } catch (error) {
        console.error('🎯 Error checking manager permissions:', error);
        setIsManager(false);
        setIsCheckingPermissions(false);
      }
    };

    checkManagerPermissions();
  }, [props.context]);

  // Rich text editor functions - safer implementation
  const contentRef = React.useRef<HTMLDivElement>(null);

  const formatText = (command: string, value?: string) => {
    try {
      // Only execute if the content ref is available and focused
      if (contentRef.current && document.activeElement === contentRef.current) {
        document.execCommand(command, false, value);
      }
    } catch (error) {
      console.warn('Format command failed:', error);
    }
  };

  const handleContentChange = (event: React.FormEvent<HTMLDivElement>) => {
    const content = event.currentTarget.innerHTML;
    setFormData({ ...formData, content });
  };

  const priorityOptions: IDropdownOption[] = [
    { key: 'High', text: '🚨 High Priority' },
    { key: 'Medium', text: '⚠️ Medium Priority' },
    { key: 'Low', text: 'ℹ️ Low Priority' }
  ];

  const audienceOptions: IDropdownOption[] = [
    { key: 'All Teams', text: '� All Teams Channels' },
    { key: 'General Channel', text: '🏢 General Channel' },
    { key: 'Leadership Team', text: '👔 Leadership Team Chat' },
    { key: 'IT Support Channel', text: '💻 IT Support Channel' },
    { key: 'Medical Staff', text: '🏥 Medical Staff Channel' },
    { key: 'Nursing Team', text: '�‍⚕️ Nursing Team Chat' },
    { key: 'Administration', text: '📋 Administration Channel' },
    { key: 'Emergency Response', text: '🚨 Emergency Response Channel' },
    { key: 'Department Heads', text: '🎯 Department Heads Chat' },
    { key: 'Custom Teams', text: '✏️ Custom Teams/Channels' }
  ];

  const expiryOptions: IDropdownOption[] = [
    { key: '1', text: '1 Day' },
    { key: '3', text: '3 Days' },
    { key: '7', text: '1 Week' },
    { key: '14', text: '2 Weeks' },
    { key: '30', text: '1 Month' }
  ];

  const handleSubmit = async () => {
    if (!currentSite) {
      setResult({ type: 'error', message: '❌ Please select a SharePoint site first' });
      return;
    }

    if (!formData.title.trim() || !formData.content.trim()) {
      setResult({ type: 'error', message: '❌ Please fill in title and content' });
      return;
    }

    setIsSubmitting(true);
    setResult({ type: 'info', message: `📤 Creating message in ${currentSiteName}...` });

    try {
      // 1. Initialize enhanced data service if not already done
      if (!enhancedDataService.getCurrentUser()) {
        await enhancedDataService.initialize(props.context, currentSite);
      } else {
        // Update site URL if changed
        enhancedDataService.setSharePointSiteUrl(currentSite);
      }
      
      console.log('🔍 DEBUG: Enhanced Data Service initialized');
      console.log('🔍 DEBUG: Current site URL:', currentSite);
      console.log('🔍 DEBUG: Current site name:', currentSiteName);
      
      if (props.context) {
        console.log('🔍 DEBUG: SPFx web URL:', props.context.pageContext?.web?.absoluteUrl);
        console.log('🔍 DEBUG: SPFx web title:', props.context.pageContext?.web?.title);
        console.log('🔍 DEBUG: SPFx user:', props.context.pageContext?.user?.displayName);
        
        // CRITICAL DEBUGGING: Let's see what URLs we're working with
        const contextSiteUrl = props.context.pageContext?.web?.absoluteUrl;
        console.log('🔍 CRITICAL DEBUG:');
        console.log('   - currentSite state:', currentSite);
        console.log('   - SPFx context site:', contextSiteUrl);
        console.log('   - Are they the same?', currentSite === contextSiteUrl);
        
        // IMPORTANT: Always use the enhanced data service
        console.log('🔧 Enhanced Data Service configured for site:', currentSite);
        
        // Check if we're trying to access a different site than the current context
        if (currentSite && contextSiteUrl && !currentSite.startsWith(contextSiteUrl) && !contextSiteUrl.startsWith(currentSite)) {
          console.warn('⚠️ CROSS-SITE ACCESS DETECTED:');
          console.warn('   Context site:', contextSiteUrl);
          console.warn('   Target site:', currentSite);
          console.warn('   This may cause 403 Forbidden errors!');
          
          setResult({ 
            type: 'error', 
            message: `❌ Cross-site access detected!\n\nContext site: ${contextSiteUrl}\nTarget site: ${currentSite}\n\n💡 You may not have permission to access the target site from this context. Try:\n1. Opening the web part directly on the target site\n2. Using the same site for both context and target\n3. Ensuring you have proper cross-site permissions` 
          });
          return;
        }
      } else {
        console.warn('⚠️ No SPFx context available - this may cause authentication issues');
        setResult({ 
          type: 'error', 
          message: '❌ No SharePoint context available!\n\nThis component requires SPFx context to access SharePoint. Make sure:\n1. The web part is added to a SharePoint page\n2. You\'re not viewing in preview mode\n3. The page has fully loaded' 
        });
        return;
      }

      const expiryDate = new Date();
      expiryDate.setDate(expiryDate.getDate() + parseInt(formData.expiryDays));

      // Determine source based on context
      const isTeamsContext = !props.context || window.location.href.includes('teams.microsoft.com');
      
      const newMessage = {
        Title: formData.title,
        MessageContent: formData.content,
        Priority: formData.priority as 'High' | 'Medium' | 'Low',
        TargetAudience: formData.targetAudience,
        ExpiryDate: expiryDate,
        Source: isTeamsContext ? 'Teams' as const : 'SharePoint' as const
      };

      console.log('📝 Creating message with data:', newMessage);
      console.log('🎯 Target site:', currentSite);
      console.log('🔗 SharePoint context site:', props.context.pageContext?.web?.absoluteUrl);
      console.log('👤 Current user:', props.context.pageContext?.user?.displayName);
      console.log('📧 User email:', props.context.pageContext?.user?.email);
      console.log('🌐 Window location:', window.location.href);
      
      // Validate that we have proper SharePoint context
      if (!props.context.pageContext?.web?.absoluteUrl) {
        setResult({ 
          type: 'error', 
          message: '❌ Invalid SharePoint context!\n\nThe web context is not available. This usually means:\n1. The component is not running in a proper SharePoint context\n2. The page hasn\'t fully loaded\n3. There\'s a permissions issue with the current site' 
        });
        return;
      }
      
      // Create message using enhanced data service
      const messageId = await enhancedDataService.createMessage(newMessage);
      console.log('✅ Message created with ID:', messageId);
      
      if (!messageId || messageId <= 0) {
        setResult({ 
          type: 'error', 
          message: '❌ Message creation failed!\n\nThe message was not created successfully. Check:\n1. SharePoint list "Important Messages" exists\n2. You have contribute permissions\n3. Required fields are properly configured\n4. Browser console for detailed error messages' 
        });
        return;
      }
      
      // 2. Get the full message for distribution
      const fullMessage = await enhancedDataService.getMessageById(messageId);
      console.log('📄 Retrieved full message:', fullMessage);

      // 3. Choose distribution method: Email or Webhook
      if (formData.useEmailIntegration) {
        // 📧 NEW: Enhanced Teams integration using Graph API
        console.log('📧 Using enhanced Teams integration...');
        
        const emailResult = await EnhancedTeamsService.distributeToAccessibleChannels(fullMessage);
        
        if (emailResult.success === 0) {
          setResult({ 
            type: 'error', 
            message: `❌ No Teams channels accessible!\n\nMessage created in SharePoint (ID: ${messageId}) but no accessible Teams channels found.\n\n💡 Make sure you have access to Teams channels or check permissions.` 
          });
        } else if (emailResult.failed === 0) {
          setResult({ 
            type: 'success', 
            message: `✅ Message created and sent to Teams!\n📊 Sent to ${emailResult.success} Teams channels\n📋 Message ID: ${messageId}` 
          });
        } else {
          setResult({ 
            type: 'success', 
            message: `⚠️ Partial success!\n📊 Sent to ${emailResult.success} channels, ${emailResult.failed} failed\n📋 Message ID: ${messageId}\n\n💡 Check Teams permissions and channel access.` 
          });
        }
        
      } else if (webhookUrls.trim()) {
        // 🔗 Enhanced Teams distribution using Graph integration
        console.log('🔗 Using enhanced Teams integration...');
        
        // Create alternative distribution methods since we removed external webhooks
        const htmlNotification = await EnhancedTeamsService.createNotification(fullMessage);
        const shareLink = await EnhancedTeamsService.createShareableLink(fullMessage);
        const copyPasteMessage = EnhancedTeamsService.createCopyPasteMessage(fullMessage);
        
        console.log('✅ Created alternative distribution content');
        
        setResult({ 
          type: 'success', 
          message: `✅ Message created with enhanced distribution options!\n\n� Message ID: ${messageId}\n🔗 Shareable link created\n📝 Copy-paste message ready\n💡 Use the dashboard to view and share the message` 
        });
      } else {
        setResult({ 
          type: 'success', 
          message: `✅ Message created in SharePoint!\nMessage ID: ${messageId}\n💡 Enable email integration or add webhook URLs to distribute to Teams` 
        });
      }

      // Reset form
      setFormData({
        title: '',
        content: '',
        priority: 'Medium',
        targetAudience: 'Teams Channel',
        expiryDays: '7',
        distributionChannels: [],
        useEmailIntegration: false
      });
      
      // Clear the rich text editor safely
      if (contentRef.current) {
        contentRef.current.innerHTML = '';
      }
      setWebhookUrls('');

      if (props.onMessageCreated) {
        props.onMessageCreated(messageId);
      }

    } catch (error) {
      console.error('❌ Error creating message:', error);
      console.error('📋 Form data was:', formData);
      console.error('🎯 Target site was:', currentSite);
      console.error('💾 Message data was:', {
        Title: formData.title,
        MessageContent: formData.content,
        Priority: formData.priority,
        TargetAudience: formData.targetAudience
      });
      
      // More specific error message
      let errorMessage = `❌ Failed to create message: ${error.message}`;
      
      if (error.message.includes('404') || error.message.includes('Not Found')) {
        errorMessage += '\n\n💡 Possible issues:\n• SharePoint list "Important Messages" may not exist\n• You may not have access to the selected site\n• The list may have a different name';
      } else if (error.message.includes('400') || error.message.includes('Bad Request')) {
        errorMessage += '\n\n💡 Possible issues:\n• Required field may be missing from SharePoint list\n• Field types may not match\n• Data validation failed';
      } else if (error.message.includes('403') || error.message.includes('Forbidden')) {
        errorMessage += '\n\n💡 Possible issues:\n• You don\'t have permission to add items to the list\n• The site may require additional permissions';
      }
      
      setResult({ type: 'error', message: errorMessage });
    } finally {
      setIsSubmitting(false);
    }
  };

  // Authentication test functions
  const runAuthTest = async () => {
    if (!props.context) {
      setResult({ type: 'error', message: '❌ No SPFx context available for authentication test' });
      return;
    }

    setResult({ type: 'info', message: '🔍 Running SharePoint authentication test...' });

    // Use the enhanced data service to test authentication
    try {
      await enhancedDataService.initialize(props.context, currentSite);
      const user = enhancedDataService.getCurrentUser();
      if (user) {
        setResult({ type: 'success', message: `✅ Authentication test passed!\nUser: ${user.spfx?.displayName || user.spfx?.email || 'Unknown'}` });
      } else {
        setResult({ type: 'error', message: '❌ Authentication test failed - could not get current user' });
      }
    } catch (error) {
      setResult({ type: 'error', message: `❌ Authentication test failed!\n${error.message}` });
    }
  };

  const testMessageCreation = async () => {
    if (!props.context) {
      setResult({ type: 'error', message: '❌ No SPFx context available for message creation test' });
      return;
    }

    setResult({ type: 'info', message: '📝 Testing message creation directly...' });

    // Use enhanced data service to test message creation
    try {
      await enhancedDataService.initialize(props.context, currentSite);
      
      const testMessage = {
        Title: 'Test Message',
        MessageContent: 'This is a test message to verify functionality.',
        Priority: 'Medium' as 'Medium',
        TargetAudience: 'Test',
        ExpiryDate: new Date(Date.now() + 24 * 60 * 60 * 1000)
      };
      
      const messageId = await enhancedDataService.createMessage(testMessage);
      setResult({ type: 'success', message: `✅ Message creation test passed!\nMessage ID: ${messageId}` });
    } catch (error) {
      setResult({ type: 'error', message: `❌ Message creation test failed!\n${error.message}` });
    }
  };

  const handleQuickTemplate = (template: 'urgent' | 'maintenance' | 'announcement' | 'routine') => {
    const templates = {
      urgent: {
        title: '🚨 Verksamhetskritisk Information',
        content: '<p><strong>Detta är verksamhetskritisk information</strong> som kräver <em>omedelbar uppmärksamhet</em>.</p><p style="color: #d73a49;">Vänligen granska och vidta nödvändiga åtgärder.</p>',
        priority: 'High',
        targetAudience: 'Teams Channel',
        expiryDays: '1',
        distributionChannels: [] as string[]
      },
      maintenance: {
        title: '🔧 Viktig information!',
        content: '<p><strong>Viktig information</strong> som berör verksamheten.</p><ul><li>Läs igenom denna information noggrant</li><li>Kontakta ansvarig vid frågor</li></ul>',
        priority: 'Medium',
        targetAudience: 'Chat Group',
        expiryDays: '3',
        distributionChannels: [] as string[]
      },
      announcement: {
        title: '📢 Notera',
        content: '<p>Information som är bra att känna till.</p><p style="color: #0366d6;"><em>Läs igenom när du har tid.</em></p>',
        priority: 'Low',
        targetAudience: 'Teams Channel',
        expiryDays: '7',
        distributionChannels: [] as string[]
      },
      routine: {
        title: '📢 Uppdaterad/Ny Rutin',
        content: '<p style="color: #0366d6;"><strong>Ny eller uppdaterad rutin</strong> har implementerats.</p><p style="color: #0366d6;"><em>Vänligen läs igenom och följ de nya riktlinjerna.</em></p><ul><li style="color: #0366d6;">Granska rutinändringarna</li><li style="color: #0366d6;">Implementera i dagligt arbete</li><li style="color: #0366d6;">Kontakta ansvarig vid frågor</li></ul>',
        priority: 'Low',
        targetAudience: 'Teams Channel',
        expiryDays: '7',
        distributionChannels: [] as string[]
      }
    };

    const template_data = templates[template];
    
    // Update all form fields with template data
    setFormData({
      title: template_data.title,
      content: template_data.content,
      priority: template_data.priority,
      targetAudience: template_data.targetAudience,
      expiryDays: template_data.expiryDays,
      distributionChannels: template_data.distributionChannels,
      useEmailIntegration: false
    });
    
    // Update the rich text editor content safely
    if (contentRef.current) {
      contentRef.current.innerHTML = template_data.content;
    }
  };

  const handleSiteSelected = (siteUrl: string, siteName: string) => {
    setCurrentSite(siteUrl);
    setCurrentSiteName(siteName);
    setResult({ type: 'info', message: `✅ Connected to ${siteName}` });
  };

  return (
    <div style={{ padding: '20px', maxWidth: '800px' }}>
      <h2>📝 Create Message from Teams</h2>
      <p>Create and distribute important messages directly from Microsoft Teams</p>
      
      {/* Debug info - only shows in console, this is to verify render is called */}
      {console.log('🎯 TeamsMessageCreator render() called - Component is rendering!')}

      {/* Permission checking state */}
      {isCheckingPermissions && (
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <Spinner size={SpinnerSize.large} label="Checking permissions..." />
          <p style={{ marginTop: '10px', color: '#666' }}>
            Verifying your manager access from SharePoint list...
          </p>
        </div>
      )}

      {/* Access denied for non-managers */}
      {!isCheckingPermissions && isManager === false && (
        <div style={{ padding: '20px', textAlign: 'center' }}>
          <MessageBar
            messageBarType={MessageBarType.blocked}
            isMultiline={true}
          >
            <h3>🔒 Access Restricted</h3>
            <p>
              <strong>Message creation is restricted to managers only.</strong>
            </p>
            <p>
              You are not currently listed as a manager in the SharePoint Managers list. 
              If you believe this is an error, please contact your administrator.
            </p>
            <div style={{ marginTop: '15px', padding: '10px', backgroundColor: '#fff3cd', borderRadius: '4px' }}>
              <strong>How manager access is determined:</strong>
              <ul style={{ textAlign: 'left', marginTop: '8px' }}>
                <li>Your email must be listed in the "Managers" SharePoint list</li>
                <li>Your entry must have "Is Active" set to "Yes"</li>
                <li>Contact HR or IT to be added to the managers list</li>
              </ul>
            </div>
          </MessageBar>
        </div>
      )}

      {/* Manager access granted - show the full interface */}
      {!isCheckingPermissions && isManager === true && (
        <>
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            dismissButtonAriaLabel="Close"
          >
            ✅ Manager access confirmed. You can create and distribute messages.
          </MessageBar>

          {/* Site Selector */}
          <SiteSelector 
            onSiteSelected={handleSiteSelected}
            currentSite={currentSite}
          />

          {/* Show form only after site is selected */}
          {currentSite && (
            <>
          {/* Quick Templates */}
          <div style={{ marginBottom: '20px', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '8px' }}>
            <h4>⚡ Quick Templates:</h4>
            <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
              <DefaultButton text="🚨 Verksamhetskritisk" onClick={() => handleQuickTemplate('urgent')} />
              <DefaultButton text="🔧 Viktig information" onClick={() => handleQuickTemplate('maintenance')} />
              <DefaultButton text="📢 Notera" onClick={() => handleQuickTemplate('announcement')} />
              <DefaultButton text="📢 Uppdaterad/Ny Rutin" onClick={() => handleQuickTemplate('routine')} />
            </div>
          </div>

      {/* Message Form */}
      <div style={{ display: 'grid', gap: '15px' }}>
        <TextField
          label="📋 Message Title"
          value={formData.title}
          onChange={(_, value) => setFormData({ ...formData, title: value || '' })}
          placeholder="Enter a clear, descriptive title"
          required
        />

        <div style={{ marginBottom: '15px' }}>
          <Label required>📝 Message Content</Label>
          
          {/* Rich Text Editor Toolbar */}
          <div style={{ 
            border: '1px solid #d0d7de', 
            borderBottom: 'none',
            padding: '8px',
            backgroundColor: '#f6f8fa',
            display: 'flex',
            gap: '4px',
            flexWrap: 'wrap'
          }}>
            {/* Font formatting */}
            <button type="button" onClick={() => formatText('bold')} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              <strong>B</strong>
            </button>
            <button type="button" onClick={() => formatText('italic')} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              <em>I</em>
            </button>
            <button type="button" onClick={() => formatText('underline')} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              <u>U</u>
            </button>
            
            {/* Font size */}
            <select onChange={(e) => formatText('fontSize', e.target.value)} style={{ padding: '4px', border: '1px solid #ccc' }}>
              <option value="">Font Size</option>
              <option value="1">Small</option>
              <option value="3">Normal</option>
              <option value="5">Large</option>
              <option value="7">Extra Large</option>
            </select>
            
            {/* Font color */}
            <input 
              type="color" 
              onChange={(e) => formatText('foreColor', e.target.value)}
              style={{ width: '30px', height: '26px', border: '1px solid #ccc' }}
              title="Text Color"
            />
            
            {/* Background color */}
            <input 
              type="color" 
              onChange={(e) => formatText('backColor', e.target.value)}
              style={{ width: '30px', height: '26px', border: '1px solid #ccc' }}
              title="Background Color"
            />
            
            {/* Lists */}
            <button type="button" onClick={() => formatText('insertUnorderedList')} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              • List
            </button>
            <button type="button" onClick={() => formatText('insertOrderedList')} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              1. List
            </button>
            
            {/* Links and tables */}
            <button type="button" onClick={() => {
              const url = prompt('Enter URL:');
              if (url) formatText('createLink', url);
            }} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              🔗 Link
            </button>
            
            <button type="button" onClick={() => {
              const tableHtml = '<table border="1" style="border-collapse: collapse; width: 100%;"><tr><td style="padding: 8px;">Cell 1</td><td style="padding: 8px;">Cell 2</td></tr><tr><td style="padding: 8px;">Cell 3</td><td style="padding: 8px;">Cell 4</td></tr></table>';
              formatText('insertHTML', tableHtml);
            }} style={{ padding: '4px 8px', border: '1px solid #ccc', background: '#fff' }}>
              📊 Table
            </button>
          </div>
          
          {/* Rich Text Content Area */}
          <div
            ref={contentRef}
            contentEditable
            onInput={handleContentChange}
            dangerouslySetInnerHTML={{ __html: formData.content }}
            dir="ltr"
            lang="sv-SE"
            style={{
              border: '1px solid #d0d7de',
              minHeight: '120px',
              padding: '12px',
              backgroundColor: '#fff',
              outline: 'none',
              fontSize: '14px',
              lineHeight: '1.5',
              fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
              direction: 'ltr',
              textAlign: 'left',
              unicodeBidi: 'embed'
            } as React.CSSProperties}
            placeholder="Enter your message content with rich formatting..."
          />
          
          <div style={{ fontSize: '12px', color: '#666', marginTop: '4px' }}>
            💡 Use the toolbar above to format text, add links, tables, and more
          </div>
        </div>

        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px' }}>
          <Dropdown
            label="⚡ Priority"
            selectedKey={formData.priority}
            onChange={(_, option) => setFormData({ ...formData, priority: option?.key as string || 'Medium' })}
            options={priorityOptions}
          />

          <Dropdown
            label="👥 Target Audience"
            selectedKey={formData.targetAudience}
            onChange={(_, option) => setFormData({ ...formData, targetAudience: option?.key as string || 'Teams Channel' })}
            options={audienceOptions}
          />

          <Dropdown
            label="📅 Expires In"
            selectedKey={formData.expiryDays}
            onChange={(_, option) => setFormData({ ...formData, expiryDays: option?.key as string || '7' })}
            options={expiryOptions}
          />
        </div>

        {/* Teams Integration Method */}
        <div style={{ marginTop: '20px', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '8px', border: '1px solid #e1e5e9' }}>
          <h4 style={{ margin: '0 0 10px 0', color: '#323130' }}>📧 Teams Integration Method</h4>
          
          <div style={{ display: 'flex', gap: '20px', marginBottom: '15px' }}>
            <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
              <input
                type="radio"
                name="teamsIntegration"
                checked={formData.useEmailIntegration}
                onChange={() => setFormData({ ...formData, useEmailIntegration: true })}
                style={{ marginRight: '8px' }}
              />
              <span>📧 <strong>Email Integration</strong> (Easy - uses SharePoint list)</span>
            </label>
            
            <label style={{ display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
              <input
                type="radio"
                name="teamsIntegration"
                checked={!formData.useEmailIntegration}
                onChange={() => setFormData({ ...formData, useEmailIntegration: false })}
                style={{ marginRight: '8px' }}
              />
              <span>🔗 <strong>Webhook Integration</strong> (Manual setup required)</span>
            </label>
          </div>

          {formData.useEmailIntegration ? (
            <div style={{ backgroundColor: '#fff', padding: '10px', borderRadius: '4px', border: '1px solid #d1d9e0' }}>
              <p style={{ margin: '0', fontSize: '14px', color: '#605e5c' }}>
                ✅ <strong>Automatic sending to configured Teams channels</strong><br/>
                📋 Channels are configured in the <a href="https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/Lists/TeamsChannels/AllItems.aspx" target="_blank" rel="noopener noreferrer">TeamsChannels SharePoint list</a><br/>
                🎯 Messages will be sent based on priority and department filters
              </p>
            </div>
          ) : (
            <div style={{ backgroundColor: '#fff', padding: '10px', borderRadius: '4px', border: '1px solid #d1d9e0' }}>
              <p style={{ margin: '0 0 10px 0', fontSize: '14px', color: '#605e5c' }}>
                🔗 <strong>Manual webhook setup required</strong><br/>
                💡 Get webhook URLs from Teams channels (Channel → ... → Connectors → Incoming Webhook)
              </p>
            </div>
          )}
        </div>

        {!formData.useEmailIntegration && (
          <TextField
            label="🔗 Teams Webhook URLs (one per line)"
            value={webhookUrls}
            onChange={(_, value) => setWebhookUrls(value || '')}
            placeholder={`https://outlook.office.com/webhook/channel1...\nhttps://outlook.office.com/webhook/channel2...`}
            multiline
            rows={3}
            description="Paste webhook URLs from Teams channels where you want to distribute this message"
          />
        )}
      </div>

      {/* Actions */}
      <div style={{ marginTop: '20px', display: 'flex', gap: '10px' }}>
        <PrimaryButton
          text="📤 Create & Distribute"
          onClick={handleSubmit}
          disabled={isSubmitting || !formData.title.trim() || !formData.content.trim()}
        />
        <DefaultButton
          text="💾 Save to SharePoint Only"
          onClick={handleSubmit}
          disabled={isSubmitting}
        />
      </div>

      {/* Result */}
      {result && (
        <MessageBar
          messageBarType={result.type === 'success' ? MessageBarType.success : 
                        result.type === 'error' ? MessageBarType.error : MessageBarType.info}
          styles={{ root: { marginTop: '20px' } }}
        >
          <pre style={{ whiteSpace: 'pre-wrap', fontFamily: 'inherit' }}>
            {result.message}
          </pre>
        </MessageBar>
      )}

      {/* Debugging Section - Only show if there are context issues */}
      {props.context && (
        <div style={{ marginTop: '20px', padding: '15px', backgroundColor: '#fff8e6', borderRadius: '8px', border: '1px solid #ffd700' }}>
          <h4>🔧 Debugging Tools</h4>
          <p>If message creation is failing, use these tests to diagnose the issue:</p>
          <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
            <DefaultButton 
              text="🔍 Test Authentication" 
              onClick={runAuthTest}
              disabled={isSubmitting}
            />
            <DefaultButton 
              text="📝 Test Message Creation" 
              onClick={testMessageCreation}
              disabled={isSubmitting}
            />
          </div>
          <div style={{ fontSize: '12px', color: '#666', marginTop: '8px' }}>
            💡 These tests will check if you have proper access to SharePoint and can create messages
          </div>
        </div>
      )}

      {/* Help Section */}
      <div style={{ marginTop: '30px', padding: '15px', backgroundColor: '#e8f4fd', borderRadius: '8px' }}>
        <h4>💡 How to Use:</h4>
        <ol>
          <li><strong>🎯 From Teams Channel:</strong> Get webhook URL (Channel → ⋯ → Connectors → Incoming Webhook)</li>
          <li><strong>📝 Create Message:</strong> Fill out the form above with your message details</li>
          <li><strong>📤 Distribute:</strong> Message goes to SharePoint + selected Teams channels</li>
          <li><strong>📊 Track:</strong> View read confirmations in the dashboard</li>
        </ol>
        
        <h4>🔄 Integration Options:</h4>
        <ul>
          <li><strong>Teams Tab:</strong> Add this as a tab in your Teams channel</li>
          <li><strong>Teams Bot:</strong> Create a bot for conversational message creation</li>
          <li><strong>Power Automate:</strong> Trigger from Teams messages or reactions</li>
          <li><strong>Teams App:</strong> Package as a full Teams application</li>
        </ul>
      </div>
            </>
          )}
        </>
      )}
    </div>
  );
};
