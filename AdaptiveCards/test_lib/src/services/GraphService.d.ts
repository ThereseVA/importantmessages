import { WebPartContext } from '@microsoft/sp-webpart-base';
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
export declare class GraphService {
    private context;
    private graphClient;
    /**
     * Initialize the Graph service with SPFx context
     */
    initialize(context: WebPartContext): Promise<void>;
    /**
     * Get current user information without requiring admin approval
     */
    getCurrentUser(): Promise<IGraphUser | null>;
    /**
     * Get user's groups - this uses delegated permissions (user's own access)
     */
    getCurrentUserGroups(): Promise<string[]>;
    /**
     * Get SharePoint sites the user has access to - using delegated permissions
     */
    getUserAccessibleSites(): Promise<IGraphSite[]>;
    /**
     * Send Teams message using delegated permissions
     */
    sendTeamsMessage(chatId: string, message: string): Promise<boolean>;
    /**
     * Check if user has access to a specific SharePoint site
     */
    checkSiteAccess(siteUrl: string): Promise<boolean>;
    /**
     * Get user photo as base64 string
     */
    getUserPhoto(): Promise<string | null>;
    /**
     * Create a SharePoint list item using Graph API (if user has permissions)
     */
    createListItem(siteId: string, listId: string, fields: any): Promise<any>;
    /**
     * Get SharePoint list items using Graph API (if user has permissions)
     */
    getListItems(siteId: string, listId: string, filter?: string): Promise<any[]>;
    /**
     * Check if the service is properly initialized
     */
    isInitialized(): boolean;
    /**
     * Get enhanced user information by combining Graph and SPFx context
     */
    getEnhancedUserInfo(): Promise<{
        graph: IGraphUser | null;
        context: any;
        groups: string[];
        photo?: string;
        isManager: boolean;
        isAdmin: boolean;
    }>;
    /**
     * Check if user has manager status based on job title or groups
     */
    private checkManagerStatus;
    /**
     * Check if user has admin status based on groups
     */
    private checkAdminStatus;
}
export declare const graphService: GraphService;
//# sourceMappingURL=GraphService.d.ts.map