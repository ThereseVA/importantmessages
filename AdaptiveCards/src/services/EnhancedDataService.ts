import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IHttpClientOptions } from '@microsoft/sp-http';
import { graphService, IGraphUser } from './GraphService';
import { ManagersListService } from './ManagersListService';

// Keep existing interfaces
export interface IMessage {
  Id: number;
  Title: string;
  MessageContent: string;
  Priority: 'High' | 'Medium' | 'Low';
  ExpiryDate?: Date;
  TargetAudience: string;
  Source?: 'Teams' | 'SharePoint' | 'Outlook';
  ReadById?: number[];
  ReadBy?: string;
  Created: Date;
  Modified: Date;
  Author: {
    Title: string;
    Email: string;
  };
}

export interface IReadAction {
  Id?: number;
  MessageId: number;
  UserId: number;
  UserEmail: string;
  UserDisplayName: string;
  ReadTimestamp: Date;
  DeviceInfo?: string;
}

export interface IEnhancedUser {
  displayName: string;
  email: string;
  graph: IGraphUser | null;
  spfx: {
    displayName: string;
    email: string;
    loginName: string;
  } | null;
  groups: string[];
  hasPhoto: boolean;
  isManager: boolean;
  isAdmin: boolean;
}

export class EnhancedDataService {
  private context: WebPartContext;
  private managersService: ManagersListService;
  private readonly MESSAGES_LIST = 'Important Messages';
  private readonly READ_ACTIONS_LIST = 'MessageReadConfirmations';
  private customSiteUrl: string = '';
  private currentUser: IEnhancedUser | null = null;

  /**
   * Initialize the service with SharePoint context and Graph service
   */
  public async initialize(context: WebPartContext, dataSourceUrl?: string): Promise<void> {
    this.context = context;
    this.managersService = new ManagersListService(context);
    
    try {
      // Initialize Graph service with proper error handling
      console.log('EnhancedDataService: Initializing Graph service...');
      await graphService.initialize(context);
      console.log('EnhancedDataService: Graph service initialized successfully');
    } catch (error) {
      console.warn('EnhancedDataService: Graph service initialization failed, continuing without Graph:', error);
    }
    
    try {
      // Get enhanced user information with fallback
      console.log('EnhancedDataService: Getting enhanced user information...');
      this.currentUser = await this.getEnhancedCurrentUser();
      console.log('EnhancedDataService: Enhanced user information retrieved');
    } catch (error) {
      console.warn('EnhancedDataService: Failed to get enhanced user info, using basic fallback:', error);
      this.currentUser = {
        displayName: context?.pageContext?.user?.displayName || 'Unknown',
        email: context?.pageContext?.user?.email || '',
        graph: null,
        spfx: context?.pageContext?.user ? {
          displayName: context.pageContext.user.displayName,
          email: context.pageContext.user.email,
          loginName: context.pageContext.user.loginName
        } : null,
        groups: [],
        hasPhoto: false,
        isManager: false,
        isAdmin: false
      };
    }
    
    // Set custom site URL if provided
    if (dataSourceUrl && dataSourceUrl.includes('sharepoint.com')) {
      const match = dataSourceUrl.match(/(https:\/\/[^\/]+\/[^\/]+\/[^\/]+)/);
      if (match) {
        this.customSiteUrl = match[1];
        console.log(`EnhancedDataService: Using custom site URL: ${this.customSiteUrl}`);
      }
    }
    
    console.log('EnhancedDataService: Initialized successfully');
  }

  /**
   * Get enhanced user information combining Graph API and SPFx context
   */
  public async getEnhancedCurrentUser(): Promise<IEnhancedUser> {
    try {
      const enhanced = await graphService.getEnhancedUserInfo();
      const userEmail = enhanced.graph?.mail || enhanced.context?.email || '';
      
      // Check if user is manager using SharePoint Managers list
      let isManager = false;
      try {
        if (this.managersService && userEmail) {
          isManager = await this.managersService.isUserManager(userEmail);
        }
      } catch (error) {
        console.warn('EnhancedDataService: Error checking manager status from SharePoint list:', error);
        // Fallback to Graph service result
        isManager = enhanced.isManager;
      }
      
      return {
        displayName: enhanced.graph?.displayName || enhanced.context?.displayName || 'Unknown',
        email: userEmail,
        graph: enhanced.graph,
        spfx: enhanced.context,
        groups: enhanced.groups,
        hasPhoto: !!enhanced.photo,
        isManager: isManager,
        isAdmin: enhanced.isAdmin
      };
    } catch (error) {
      console.warn('EnhancedDataService: Error getting enhanced user info, using SPFx fallback:', error);
      
      // Fallback to SPFx context only with SharePoint manager check
      const email = this.context?.pageContext?.user?.email || '';
      let isManager = false;
      
      try {
        if (this.managersService && email) {
          isManager = await this.managersService.isUserManager(email);
        }
      } catch (error) {
        console.warn('EnhancedDataService: Error checking manager status in fallback:', error);
      }
      
      return {
        displayName: this.context?.pageContext?.user?.displayName || 'Unknown',
        email: email,
        graph: null,
        spfx: this.context?.pageContext?.user ? {
          displayName: this.context.pageContext.user.displayName,
          email: this.context.pageContext.user.email,
          loginName: this.context.pageContext.user.loginName
        } : null,
        groups: [],
        hasPhoto: false,
        isManager: isManager,
        isAdmin: false
      };
    }
  }

  /**
   * Check if the current user is a manager according to SharePoint Managers list
   */
  public async isCurrentUserManager(): Promise<boolean> {
    try {
      if (!this.managersService) {
        console.warn('EnhancedDataService: ManagersListService not initialized');
        return false;
      }
      
      const userEmail = this.context?.pageContext?.user?.email;
      if (!userEmail) {
        console.warn('EnhancedDataService: No user email available');
        return false;
      }
      
      return await this.managersService.isUserManager(userEmail);
    } catch (error) {
      console.error('EnhancedDataService: Error checking manager status:', error);
      return false;
    }
  }

  /**
   * Check if a specific user is a manager
   */
  public async isUserManager(userEmail: string): Promise<boolean> {
    try {
      if (!this.managersService) {
        console.warn('EnhancedDataService: ManagersListService not initialized');
        return false;
      }
      
      return await this.managersService.isUserManager(userEmail);
    } catch (error) {
      console.error('EnhancedDataService: Error checking manager status for user:', userEmail, error);
      return false;
    }
  }

  /**
   * Get manager details for the current user
   */
  public async getCurrentUserManagerDetails() {
    try {
      if (!this.managersService) {
        console.warn('EnhancedDataService: ManagersListService not initialized');
        return null;
      }
      
      const userEmail = this.context?.pageContext?.user?.email;
      if (!userEmail) {
        console.warn('EnhancedDataService: No user email available');
        return null;
      }
      
      return await this.managersService.getManagerDetails(userEmail);
    } catch (error) {
      console.error('EnhancedDataService: Error getting manager details:', error);
      return null;
    }
  }

  /**
   * Get all active managers from SharePoint list
   */
  public async getAllManagers() {
    try {
      if (!this.managersService) {
        console.warn('EnhancedDataService: ManagersListService not initialized');
        return [];
      }
      
      return await this.managersService.getActiveManagers();
    } catch (error) {
      console.error('EnhancedDataService: Error getting all managers:', error);
      return [];
    }
  }

  /**
   * Set a custom SharePoint site URL
   */
  public setSharePointSiteUrl(siteUrl: string): void {
    const normalizedUrl = siteUrl.replace(/\/$/, '');
    console.log(`EnhancedDataService: Setting custom SharePoint site URL: ${normalizedUrl}`);
    this.customSiteUrl = normalizedUrl;
  }

  /**
   * Get the current site URL for API calls
   */
  public getCurrentSiteUrl(): string {
    let siteUrl: string;
    
    if (this.isTeamsContext()) {
      if (this.customSiteUrl) {
        siteUrl = this.customSiteUrl;
      } else {
        const teamsSiteUrl = this.getSharePointSiteFromTeamsContext();
        if (teamsSiteUrl) {
          siteUrl = teamsSiteUrl;
        } else {
          console.warn('EnhancedDataService: Teams context detected but no SharePoint site configured.');
          siteUrl = this.context?.pageContext?.web?.absoluteUrl || 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken';
        }
      }
    } else {
      siteUrl = this.customSiteUrl || this.context?.pageContext?.web?.absoluteUrl || 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken';
    }
    
    return siteUrl.replace(/\/$/, '');
  }

  /**
   * Check if we're running in Teams context
   */
  private isTeamsContext(): boolean {
    const url = window.location.href;
    const isTeamsUrl = url.includes('teams.microsoft.com') || url.includes('teams.office.com');
    
    let hasTeamsContext = false;
    try {
      hasTeamsContext = this.context?.sdks?.microsoftTeams?.context !== undefined;
    } catch (error) {
      hasTeamsContext = false;
    }
    
    return isTeamsUrl || hasTeamsContext;
  }

  /**
   * Try to get SharePoint site URL from Teams context
   */
  private getSharePointSiteFromTeamsContext(): string | null {
    try {
      if (this.context?.sdks?.microsoftTeams?.context) {
        const teamsContext = this.context.sdks.microsoftTeams.context;
        
        if (teamsContext.sharepoint?.serverRelativeUrl) {
          const currentUrl = this.context?.pageContext?.web?.absoluteUrl;
          if (currentUrl) {
            const tenant = currentUrl.split('/')[2];
            return `https://${tenant}${teamsContext.sharepoint.serverRelativeUrl}`;
          }
        }
        
        if (teamsContext.teamSiteUrl) {
          return teamsContext.teamSiteUrl;
        }
        
        if (teamsContext.sharepoint?.webAbsoluteUrl) {
          return teamsContext.sharepoint.webAbsoluteUrl;
        }
      }
      
      return null;
    } catch (error) {
      console.warn('EnhancedDataService: Error getting SharePoint site from Teams context:', error);
      return null;
    }
  }

  /**
   * Get all active messages with enhanced user filtering
   */
  public async getActiveMessages(): Promise<IMessage[]> {
    try {
      const restUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/items` +
        `?$select=Id,Title,MessageContent,Priority,TargetAudience,ReadBy,Created,Modified` +
        `&$orderby=Priority desc,Created desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      const messages = data.value.map(this.mapToMessage);
      
      // Filter messages based on user's groups and role
      return this.filterMessagesForCurrentUser(messages);
    } catch (error) {
      console.error('EnhancedDataService: Error fetching active messages:', error);
      return this.getMockMessages();
    }
  }

  /**
   * Enhanced message filtering based on user's Graph groups and properties
   */
  private filterMessagesForCurrentUser(messages: IMessage[]): IMessage[] {
    if (!this.currentUser) {
      return messages; // Return all if no user context
    }

    const userGroups = this.currentUser.groups || [];
    const userEmail = this.currentUser.spfx?.email || this.currentUser.graph?.mail || '';
    const userDepartment = this.currentUser.graph?.department || '';
    const userJobTitle = this.currentUser.graph?.jobTitle || '';

    return messages.filter(message => {
      const targetAudience = message.TargetAudience || '';
      
      // Always show messages for all users
      if (targetAudience === 'All Users' || targetAudience === 'Alla Medarbetare') {
        return true;
      }
      
      // Check if user's groups match target audience
      const matchesGroup = userGroups.some(group => 
        targetAudience.toLowerCase().includes(group.toLowerCase())
      );
      
      // Check if user's department matches
      const matchesDepartment = userDepartment && 
        targetAudience.toLowerCase().includes(userDepartment.toLowerCase());
      
      // Check if user's job title matches
      const matchesJobTitle = userJobTitle && 
        targetAudience.toLowerCase().includes(userJobTitle.toLowerCase());
      
      // Check if user's email is specifically mentioned
      const matchesEmail = targetAudience.toLowerCase().includes(userEmail.toLowerCase());
      
      console.log(`EnhancedDataService: Message "${message.Title}" - Target: ${targetAudience}, User Groups: [${userGroups.join(', ')}], Match: ${matchesGroup || matchesDepartment || matchesJobTitle || matchesEmail}`);
      
      return matchesGroup || matchesDepartment || matchesJobTitle || matchesEmail;
    });
  }

  /**
   * Get messages for current user with enhanced targeting
   */
  public async getMessagesForCurrentUser(): Promise<IMessage[]> {
    try {
      // Ensure we have current user information
      if (!this.currentUser) {
        this.currentUser = await this.getEnhancedCurrentUser();
      }
      
      // Check if we're in development mode
      if (!this.context || !this.context.pageContext || !this.context.pageContext.web) {
        console.log('EnhancedDataService: Development mode detected - returning mock data');
        return this.getMockMessages();
      }
      
      const siteUrl = this.getCurrentSiteUrl();
      if (!siteUrl) {
        console.warn('EnhancedDataService: No site URL available - returning mock data');
        return this.getMockMessages();
      }
      
      const allMessages = await this.getMessagesWithProgressiveQuerying(siteUrl);
      if (!allMessages || allMessages.length === 0) {
        console.warn('EnhancedDataService: No messages returned - returning mock data');
        return this.getMockMessages();
      }
      
      const filteredMessages = this.filterMessagesForCurrentUser(allMessages);
      console.log(`EnhancedDataService: Found ${allMessages.length} total messages, ${filteredMessages.length} for current user`);
      
      return filteredMessages;
      
    } catch (error) {
      console.error('EnhancedDataService: Error fetching messages for current user:', error);
      return this.getMockMessages();
    }
  }

  /**
   * Mark message as read with enhanced user tracking
   */
  public async markMessageAsRead(messageId: number): Promise<void> {
    try {
      const currentUser = this.currentUser?.spfx || this.context.pageContext.user;
      if (!currentUser) {
        throw new Error('No user context available');
      }
      
      // Check if already marked as read
      const alreadyRead = await this.hasUserReadMessage(messageId);
      if (alreadyRead) {
        console.log(`Message ${messageId} already marked as read by user ${currentUser.email}`);
        return;
      }

      // Create read action record with enhanced user info
      const readAction = {
        Title: `Read action for message ${messageId}`,
        MessageId: messageId,
        UserId: parseInt(currentUser.loginName.split('|')[2] || '0'),
        UserEmail: currentUser.email,
        UserDisplayName: currentUser.displayName,
        ReadTimestamp: new Date().toISOString(),
        DeviceInfo: this.getDeviceInfo(),
        UserDepartment: this.currentUser?.graph?.department || '',
        UserJobTitle: this.currentUser?.graph?.jobTitle || ''
      };

      // Add to read actions list
      const restUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('${this.READ_ACTIONS_LIST}')/items`;
      
      const spOpts: IHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'X-RequestDigest': await this.getRequestDigest()
        },
        body: JSON.stringify(readAction)
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        restUrl,
        SPHttpClient.configurations.v1,
        spOpts
      );

      if (!response.ok) {
        throw new Error(`Failed to create read action: ${response.status} ${response.statusText}`);
      }

      // Update the ReadBy field in the main message
      await this.updateMessageReadBy(messageId, currentUser.email);

      console.log(`Message ${messageId} marked as read by ${currentUser.email}`);
    } catch (error) {
      console.error(`EnhancedDataService: Error marking message ${messageId} as read:`, error);
      throw new Error('Failed to mark message as read');
    }
  }

  /**
   * Check if current user has read a specific message
   */
  public async hasUserReadMessage(messageId: number): Promise<boolean> {
    try {
      const currentUser = this.currentUser?.spfx || this.context.pageContext.user;
      if (!currentUser) {
        return false;
      }
      
      const restUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('${this.READ_ACTIONS_LIST}')/items` +
        `?$filter=MessageId eq ${messageId} and UserEmail eq '${currentUser.email}'` +
        `&$top=1`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const data = await response.json();
      return data.value.length > 0;
    } catch (error) {
      console.error(`EnhancedDataService: Error checking read status for message ${messageId}:`, error);
      return false;
    }
  }

  /**
   * Get current user information for display
   */
  public getCurrentUser(): IEnhancedUser | null {
    return this.currentUser;
  }

  /**
   * Create a new message (admin function)
   */
  public async createMessage(message: Partial<IMessage>): Promise<number> {
    try {
      // Start with basic required fields that should always exist
      const newMessage: any = {
        Title: message.Title,
      };

      // Add optional fields only if they have values
      if (message.MessageContent) {
        newMessage.MessageContent = message.MessageContent;
      }
      
      if (message.Priority) {
        newMessage.Priority = message.Priority;
      }
      
      if (message.TargetAudience) {
        newMessage.TargetAudience = message.TargetAudience;
      }

      // Add Source field if provided
      if (message.Source) {
        newMessage.Source = message.Source;
      }

      console.log('EnhancedDataService: Creating message with data:', newMessage);
      console.log('EnhancedDataService: Target site URL:', this.getCurrentSiteUrl());

      const restUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('Important Messages')/items`;
      console.log('EnhancedDataService: REST API URL:', restUrl);
      
      const spOpts: IHttpClientOptions = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': '',
          'X-RequestDigest': await this.getRequestDigest()
        },
        body: JSON.stringify(newMessage)
      };

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        restUrl,
        SPHttpClient.configurations.v1,
        spOpts
      );

      if (!response.ok) {
        console.error('SharePoint API Request Failed:');
        console.error('- Status:', response.status, response.statusText);
        console.error('- URL:', restUrl);
        console.error('- Data sent:', newMessage);
        
        let errorDetails = `${response.status} ${response.statusText}`;
        try {
          const errorBody = await response.text();
          console.error('SharePoint API Error Details:', errorBody);
          errorDetails += ` - ${errorBody}`;
        } catch (e) {
          console.error('Could not read error response body:', e);
        }
        throw new Error(`Failed to create message: ${errorDetails}`);
      }

      const data = await response.json();
      console.log(`EnhancedDataService: Created new message with ID: ${data.Id}`);
      return data.Id;
    } catch (error) {
      console.error('EnhancedDataService: Error creating message:', error);
      const errorMsg = error instanceof Error ? error.message : 'Failed to create message';
      throw new Error(errorMsg);
    }
  }

  /**
   * Get a specific message by ID
   */
  public async getMessageById(messageId: number): Promise<IMessage> {
    try {
      const restUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/items(${messageId})` +
        `?$select=Id,Title,MessageContent,Priority,TargetAudience,ReadBy,Created,Modified`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      return this.mapToMessage(data);
    } catch (error) {
      console.error(`EnhancedDataService: Error fetching message ${messageId}:`, error);
      throw new Error(`Failed to fetch message with ID ${messageId}`);
    }
  }

  /**
   * Check if user has access to specific functionality based on Graph groups
   */
  public hasUserRole(role: string): boolean {
    if (!this.currentUser) {
      return false;
    }
    
    const userGroups = this.currentUser.groups || [];
    const userJobTitle = this.currentUser.graph?.jobTitle || '';
    
    // Check if user is in specific groups or has specific job titles
    switch (role.toLowerCase()) {
      case 'admin':
      case 'administrator':
        return userGroups.some(group => 
          group.toLowerCase().includes('admin') || 
          group.toLowerCase().includes('administrator')
        ) || userJobTitle.toLowerCase().includes('admin');
        
      case 'manager':
        return userGroups.some(group => 
          group.toLowerCase().includes('manager') || 
          group.toLowerCase().includes('lead')
        ) || userJobTitle.toLowerCase().includes('manager');
        
      case 'hr':
        return userGroups.some(group => 
          group.toLowerCase().includes('hr') || 
          group.toLowerCase().includes('human')
        ) || userJobTitle.toLowerCase().includes('hr');
        
      default:
        return userGroups.some(group => 
          group.toLowerCase().includes(role.toLowerCase())
        );
    }
  }

  // Include all the private helper methods from the original DataService
  private async getMessagesWithProgressiveQuerying(siteUrl: string): Promise<IMessage[]> {
    console.log('EnhancedDataService: Starting progressive column querying...');
    
    try {
      const availableColumns = await this.getAvailableColumns(siteUrl);
      console.log('EnhancedDataService: Available columns detected:', availableColumns);
      
      const basicQuery = await this.testBasicQuery(siteUrl);
      if (basicQuery.success) {
        console.log('EnhancedDataService: Basic query successful, trying enhanced query...');
        const enhancedQuery = await this.testEnhancedQuery(siteUrl, availableColumns);
        if (enhancedQuery.success) {
          return enhancedQuery.data;
        } else {
          console.log('EnhancedDataService: Enhanced query failed, using basic data');
          return basicQuery.data;
        }
      } else {
        console.log('EnhancedDataService: Basic query failed, using mock data');
        return this.getMockMessages();
      }
      
    } catch (error) {
      console.error('EnhancedDataService: Progressive querying failed:', error);
      return this.getMockMessages();
    }
  }

  private async testBasicQuery(siteUrl: string): Promise<{ success: boolean; data: IMessage[] }> {
    try {
      const restUrl = `${siteUrl}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/items` +
        `?$select=Id,Title,Created,Modified` +
        `&$top=10` +
        `&$orderby=Created desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return { success: false, data: [] };
      }

      const data = await response.json();
      const messages = data.value.map((item: any) => this.mapToMessageBasic(item));
      return { success: true, data: messages };
      
    } catch (error) {
      console.error('EnhancedDataService: Basic query error:', error);
      return { success: false, data: [] };
    }
  }

  private async testEnhancedQuery(siteUrl: string, availableColumns: string[]): Promise<{ success: boolean; data: IMessage[] }> {
    try {
      const safeColumns = this.buildConservativeColumnQuery(availableColumns);
      
      const restUrl = `${siteUrl}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/items` +
        `?$select=${safeColumns.select}` +
        (safeColumns.expand ? `&$expand=${safeColumns.expand}` : '') +
        `&$top=50` +
        `&$orderby=Created desc`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        restUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return { success: false, data: [] };
      }

      const data = await response.json();
      const messages = data.value.map((item: any) => this.mapToMessageWithAvailableColumns(item, availableColumns));
      return { success: true, data: messages };
      
    } catch (error) {
      console.error('EnhancedDataService: Enhanced query error:', error);
      return { success: false, data: [] };
    }
  }

  private buildConservativeColumnQuery(availableColumns: string[]): { select: string; expand: string } {
    const safeColumns = ['Id', 'Title', 'Created', 'Modified'];
    
    const criticalFields = ['MessageContent', 'Priority', 'TargetAudience', 'ReadBy'];
    for (const field of criticalFields) {
      if (availableColumns.indexOf(field) !== -1) {
        safeColumns.push(field);
      }
    }
    
    const contentFields = ['Body', 'Description'];
    for (const field of contentFields) {
      if (availableColumns.indexOf(field) !== -1 && safeColumns.indexOf('MessageContent') === -1) {
        safeColumns.push(field);
        break;
      }
    }
    
    let expandAuthor = false;
    if (availableColumns.indexOf('Author') !== -1) {
      safeColumns.push('Author/Title');
      expandAuthor = true;
    }
    
    return {
      select: safeColumns.join(','),
      expand: expandAuthor ? 'Author' : ''
    };
  }

  private async getAvailableColumns(siteUrl: string): Promise<string[]> {
    try {
      const listInfoUrl = `${siteUrl}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false`;
      
      const response = await this.context.spHttpClient.get(
        listInfoUrl,
        SPHttpClient.configurations.v1
      );
      
      if (response.ok) {
        const listInfo = await response.json();
        const columns = listInfo.value.map((field: any) => field.InternalName);
        return columns;
      } else {
        return ['Id', 'Title', 'Created', 'Modified'];
      }
    } catch (error) {
      console.warn('EnhancedDataService: Error fetching column schema:', error);
      return ['Id', 'Title', 'Created', 'Modified'];
    }
  }

  private mapToMessageBasic(item: any): IMessage {
    return {
      Id: item.Id || 0,
      Title: item.Title || 'Untitled Message',
      MessageContent: 'Click to view message details',
      Priority: 'Medium' as 'High' | 'Medium' | 'Low',
      TargetAudience: 'All Users',
      ReadBy: '',
      ReadById: [],
      Created: new Date(item.Created),
      Modified: new Date(item.Modified),
      Author: {
        Title: 'Unknown',
        Email: ''
      }
    };
  }

  private mapToMessageWithAvailableColumns(item: any, availableColumns: string[]): IMessage {
    const getFieldValue = (fieldNames: string[], defaultValue: any = null) => {
      for (const fieldName of fieldNames) {
        if (item[fieldName] !== undefined && item[fieldName] !== null) {
          return item[fieldName];
        }
      }
      return defaultValue;
    };
    
    return {
      Id: item.Id || 0,
      Title: item.Title || 'Untitled Message',
      MessageContent: getFieldValue(['MessageContent', 'Body', 'Description', 'Content'], 'No content available'),
      Priority: getFieldValue(['Priority', 'Importance', 'Level'], 'Medium') as 'High' | 'Medium' | 'Low',
      TargetAudience: getFieldValue(['TargetAudience', 'Audience', 'Group'], 'All Users'),
      ReadBy: getFieldValue(['ReadBy', 'ReadStatus'], ''),
      ReadById: getFieldValue(['ReadBy', 'ReadStatus']) ? 
        getFieldValue(['ReadBy', 'ReadStatus']).split(';').map((email: string) => email.trim()).filter((email: string) => email) : [],
      Created: new Date(item.Created),
      Modified: new Date(item.Modified),
      Author: {
        Title: item.Author?.Title || 'Unknown',
        Email: item.Author?.Email || ''
      }
    };
  }

  private mapToMessage = (item: any): IMessage => {
    return {
      Id: item.Id,
      Title: item.Title,
      MessageContent: item.MessageContent,
      Priority: item.Priority,
      ExpiryDate: item.ExpiryDate ? new Date(item.ExpiryDate) : undefined,
      TargetAudience: item.TargetAudience,
      ReadBy: item.ReadBy,
      ReadById: item.ReadBy ? item.ReadBy.split(';').map((email: string) => email.trim()).filter((email: string) => email) : [],
      Created: new Date(item.Created),
      Modified: new Date(item.Modified),
      Author: {
        Title: 'System User',
        Email: 'system@company.com'
      }
    };
  };

  private async getRequestDigest(): Promise<string> {
    try {
      if (!this.context || !this.context.spHttpClient) {
        throw new Error('EnhancedDataService context not initialized.');
      }

      const restUrl = `${this.getCurrentSiteUrl()}/_api/contextinfo`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        restUrl,
        SPHttpClient.configurations.v1,
        {}
      );

      if (!response.ok) {
        throw new Error(`Failed to get request digest: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      return data.FormDigestValue;
    } catch (error) {
      console.error('EnhancedDataService: Error getting request digest:', error);
      throw error;
    }
  }

  private async updateMessageReadBy(messageId: number, userEmail: string): Promise<void> {
    try {
      const getUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/items(${messageId})` +
        `?$select=ReadBy`;

      const getResponse: SPHttpClientResponse = await this.context.spHttpClient.get(
        getUrl,
        SPHttpClient.configurations.v1
      );

      if (!getResponse.ok) {
        console.warn('Could not fetch current ReadBy value');
        return;
      }

      const currentData = await getResponse.json();
      const currentReadBy = currentData.ReadBy || '';
      const emails = currentReadBy ? currentReadBy.split(';').filter((email: string) => email.trim()) : [];
      
      if (!emails.includes(userEmail)) {
        emails.push(userEmail);
        const updatedReadBy = emails.join(';');

        const updateUrl = `${this.getCurrentSiteUrl()}/_api/web/lists/getByTitle('${this.MESSAGES_LIST}')/items(${messageId})`;
        
        const spOpts: IHttpClientOptions = {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'X-RequestDigest': await this.getRequestDigest()
          },
          body: JSON.stringify({
            ReadBy: updatedReadBy
          })
        };

        await this.context.spHttpClient.post(
          updateUrl,
          SPHttpClient.configurations.v1,
          spOpts
        );
      }
    } catch (error) {
      console.error(`Error updating ReadBy field for message ${messageId}:`, error);
    }
  }

  private getDeviceInfo(): string {
    const userAgent = navigator.userAgent;
    const platform = navigator.platform;
    return `${platform} - ${userAgent.substring(0, 100)}`;
  }

  private getMockMessages(): IMessage[] {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const nextWeek = new Date();
    nextWeek.setDate(nextWeek.getDate() + 7);

    return [
      {
        Id: 1,
        Title: "🚀 New Feature Release - Enhanced with Graph Integration",
        MessageContent: "We're excited to announce the release of our new dashboard features with Microsoft Graph integration! No admin approval required for basic features.",
        Priority: "High" as const,
        TargetAudience: "All Users",
        ReadBy: "",
        Created: new Date(),
        Modified: new Date(),
        Author: {
          Title: "System Administrator",
          Email: "admin@company.com"
        }
      },
      {
        Id: 2,
        Title: "📅 Maintenance Window Scheduled",
        MessageContent: "Scheduled maintenance will occur this weekend from 2 AM to 6 AM. The system will be temporarily unavailable during this time.",
        Priority: "Medium" as const,
        TargetAudience: "All Users",
        ReadBy: "",
        Created: new Date(),
        Modified: new Date(),
        Author: {
          Title: "IT Team",
          Email: "it@company.com"
        }
      },
      {
        Id: 3,
        Title: "📊 Dashboard Tutorial Available - Graph Enhanced",
        MessageContent: "New to the dashboard? Check out our comprehensive tutorial featuring Microsoft Graph integration for enhanced user experience without requiring admin permissions.",
        Priority: "Low" as const,
        TargetAudience: "New Users",
        ReadBy: "",
        Created: new Date(),
        Modified: new Date(),
        Author: {
          Title: "Training Team",
          Email: "training@company.com"
        }
      }
    ];
  }
}

// Export singleton instance
export const enhancedDataService = new EnhancedDataService();
