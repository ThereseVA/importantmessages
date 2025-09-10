import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Client } from '@microsoft/microsoft-graph-client';
import { User, Site } from '@microsoft/microsoft-graph-types';

export interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
}

export interface IGraphSite {
  id: string;
  displayName: string;
  webUrl: string;
  siteCollection?: {
    hostname: string;
  };
}

export class GraphService {
  private context: WebPartContext;
  private graphClient: Client | null = null;

  /**
   * Initialize the Graph service with SPFx context
   */
  public async initialize(context: WebPartContext): Promise<void> {
    this.context = context;
    
    try {
      console.log('GraphService: Initializing with SPFx context');
      // We'll use MSGraphClientV3 directly instead of creating a custom client
      // This avoids authentication issues and uses SPFx's built-in Graph access
      console.log('GraphService: Initialized successfully');
    } catch (error) {
      console.error('GraphService: Failed to initialize:', error);
      throw error;
    }
  }

  /**
   * Get current user information without requiring admin approval
   */
  public async getCurrentUser(): Promise<IGraphUser | null> {
    try {
      if (!this.context) {
        console.warn('GraphService: Context not initialized');
        return null;
      }

      // Use SPFx's built-in Graph client factory which handles authentication
      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Get current user - this is available without admin approval
      const user: User = await msGraphClient.api('/me').get();
      
      return {
        id: user.id || '',
        displayName: user.displayName || '',
        mail: user.mail || user.userPrincipalName || '',
        userPrincipalName: user.userPrincipalName || '',
        jobTitle: user.jobTitle,
        department: user.department
      };
    } catch (error) {
      console.warn('GraphService: Error getting current user, falling back to SPFx context:', error);
      
      // Fallback to SPFx context user info
      if (this.context?.pageContext?.user) {
        const contextUser = this.context.pageContext.user;
        return {
          id: contextUser.loginName,
          displayName: contextUser.displayName,
          mail: contextUser.email,
          userPrincipalName: contextUser.loginName,
          jobTitle: undefined,
          department: undefined
        };
      }
      
      return null;
    }
  }

  /**
   * Get user's groups - this uses delegated permissions (user's own access)
   */
  public async getCurrentUserGroups(): Promise<string[]> {
    try {
      if (!this.context) {
        console.warn('GraphService: Context not initialized');
        return [];
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Get user's groups - this works with delegated permissions
      const response = await msGraphClient.api('/me/memberOf').get();
      
      const groups = response.value || [];
      return groups
        .filter((group: any) => group['@odata.type'] === '#microsoft.graph.group')
        .map((group: any) => group.displayName || group.mailNickname || group.id);
        
    } catch (error) {
      console.warn('GraphService: Error getting user groups, returning empty array:', error);
      return [];
    }
  }

  /**
   * Get SharePoint sites the user has access to - using delegated permissions
   */
  public async getUserAccessibleSites(): Promise<IGraphSite[]> {
    try {
      if (!this.context) {
        console.warn('GraphService: Context not initialized');
        return [];
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Get sites the current user has access to - delegated permission
      const response = await msGraphClient.api('/me/followedSites').get();
      
      const sites = response.value || [];
      return sites.map((site: Site) => ({
        id: site.id || '',
        displayName: site.displayName || '',
        webUrl: site.webUrl || '',
        siteCollection: site.siteCollection
      }));
      
    } catch (error) {
      console.error('GraphService: Error getting user sites:', error);
      return [];
    }
  }

  /**
   * Send Teams message using delegated permissions
   */
  public async sendTeamsMessage(chatId: string, message: string): Promise<boolean> {
    try {
      if (!this.context) {
        console.warn('GraphService: Context not initialized');
        return false;
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Send a message to a Teams chat - requires user to be a member
      await msGraphClient.api(`/chats/${chatId}/messages`).post({
        body: {
          content: message,
          contentType: 'text'
        }
      });
      
      return true;
    } catch (error) {
      console.error('GraphService: Error sending Teams message:', error);
      return false;
    }
  }

  /**
   * Check if user has access to a specific SharePoint site
   */
  public async checkSiteAccess(siteUrl: string): Promise<boolean> {
    try {
      if (!this.context) {
        return false;
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Extract hostname and path from URL
      const url = new URL(siteUrl);
      const hostname = url.hostname;
      const sitePath = url.pathname;
      
      // Try to get site by URL - if successful, user has access
      const site = await msGraphClient.api(`/sites/${hostname}:${sitePath}`).get();
      
      return !!site;
    } catch (error) {
      console.warn('GraphService: User does not have access to site:', siteUrl);
      return false;
    }
  }

  /**
   * Get user photo as base64 string
   */
  public async getUserPhoto(): Promise<string | null> {
    try {
      if (!this.context) {
        return null;
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Get user's photo - available with delegated permissions
      const photoResponse = await msGraphClient.api('/me/photo/$value').get();
      
      // Convert to base64
      const arrayBuffer = await photoResponse.arrayBuffer();
      const uint8Array = new Uint8Array(arrayBuffer);
      const base64 = btoa(String.fromCharCode.apply(null, Array.from(uint8Array)));
      
      return `data:image/jpeg;base64,${base64}`;
    } catch (error) {
      console.warn('GraphService: Could not get user photo:', error);
      return null;
    }
  }

  /**
   * Create a SharePoint list item using Graph API (if user has permissions)
   */
  public async createListItem(siteId: string, listId: string, fields: any): Promise<any> {
    try {
      if (!this.context) {
        throw new Error('GraphService not initialized');
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      // Create list item - user must have write permissions to the list
      const response = await msGraphClient.api(`/sites/${siteId}/lists/${listId}/items`).post({
        fields: fields
      });
      
      return response;
    } catch (error) {
      console.error('GraphService: Error creating list item:', error);
      throw error;
    }
  }

  /**
   * Get SharePoint list items using Graph API (if user has permissions)
   */
  public async getListItems(siteId: string, listId: string, filter?: string): Promise<any[]> {
    try {
      if (!this.context) {
        return [];
      }

      const msGraphClientFactory = this.context.msGraphClientFactory;
      const msGraphClient: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
      
      let apiCall = msGraphClient.api(`/sites/${siteId}/lists/${listId}/items`).expand('fields');
      
      if (filter) {
        apiCall = apiCall.filter(filter);
      }
      
      const response = await apiCall.get();
      return response.value || [];
    } catch (error) {
      console.error('GraphService: Error getting list items:', error);
      return [];
    }
  }

  /**
   * Check if the service is properly initialized
   */
  public isInitialized(): boolean {
    return !!this.context;
  }

  /**
   * Get enhanced user information by combining Graph and SPFx context
   */
  public async getEnhancedUserInfo(): Promise<{
    graph: IGraphUser | null;
    context: any;
    groups: string[];
    photo?: string;
    isManager: boolean;
    isAdmin: boolean;
  }> {
    const [graphUser, groups, photo] = await Promise.all([
      this.getCurrentUser(),
      this.getCurrentUserGroups(),
      this.getUserPhoto().catch(() => null)
    ]);

    // Check for manager/admin status based on groups or job title
    const isManager = this.checkManagerStatus(graphUser, groups);
    const isAdmin = this.checkAdminStatus(groups);

    return {
      graph: graphUser,
      context: this.context?.pageContext?.user ? {
        displayName: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email,
        loginName: this.context.pageContext.user.loginName
      } : null,
      groups,
      photo: photo || undefined,
      isManager,
      isAdmin
    };
  }

  /**
   * Check if user has manager status based on job title or groups
   */
  private checkManagerStatus(user: IGraphUser | null, groups: string[]): boolean {
    if (!user) return false;

    // Check job title for manager keywords
    const managerTitles = ['manager', 'director', 'supervisor', 'lead', 'head'];
    const jobTitle = user.jobTitle?.toLowerCase() || '';
    if (managerTitles.some(title => jobTitle.includes(title))) {
      return true;
    }

    // Check groups for manager groups (customize as needed)
    const managerGroups = ['managers', 'leadership', 'directors'];
    return groups.some(group => 
      managerGroups.some(managerGroup => 
        group.toLowerCase().includes(managerGroup)
      )
    );
  }

  /**
   * Check if user has admin status based on groups
   */
  private checkAdminStatus(groups: string[]): boolean {
    const adminGroups = ['administrators', 'admin', 'global administrators'];
    return groups.some(group => 
      adminGroups.some(adminGroup => 
        group.toLowerCase().includes(adminGroup)
      )
    );
  }
}

// Export singleton instance
export const graphService = new GraphService();
