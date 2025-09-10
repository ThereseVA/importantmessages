import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IGraphUser } from './GraphService';
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
export declare class EnhancedDataService {
    private context;
    private managersService;
    private readonly MESSAGES_LIST;
    private readonly READ_ACTIONS_LIST;
    private customSiteUrl;
    private currentUser;
    /**
     * Initialize the service with SharePoint context and Graph service
     */
    initialize(context: WebPartContext, dataSourceUrl?: string): Promise<void>;
    /**
     * Get enhanced user information combining Graph API and SPFx context
     */
    getEnhancedCurrentUser(): Promise<IEnhancedUser>;
    /**
     * Check if the current user is a manager according to SharePoint Managers list
     */
    isCurrentUserManager(): Promise<boolean>;
    /**
     * Check if a specific user is a manager
     */
    isUserManager(userEmail: string): Promise<boolean>;
    /**
     * Get manager details for the current user
     */
    getCurrentUserManagerDetails(): Promise<import("./ManagersListService").IManager>;
    /**
     * Get all active managers from SharePoint list
     */
    getAllManagers(): Promise<import("./ManagersListService").IManager[]>;
    /**
     * Set a custom SharePoint site URL
     */
    setSharePointSiteUrl(siteUrl: string): void;
    /**
     * Get the current site URL for API calls
     */
    getCurrentSiteUrl(): string;
    /**
     * Check if we're running in Teams context
     */
    private isTeamsContext;
    /**
     * Try to get SharePoint site URL from Teams context
     */
    private getSharePointSiteFromTeamsContext;
    /**
     * Get all active messages with enhanced user filtering
     */
    getActiveMessages(): Promise<IMessage[]>;
    /**
     * Enhanced message filtering based on user's Graph groups and properties
     */
    private filterMessagesForCurrentUser;
    /**
     * Get messages for current user with enhanced targeting
     */
    getMessagesForCurrentUser(): Promise<IMessage[]>;
    /**
     * Mark message as read with enhanced user tracking
     */
    markMessageAsRead(messageId: number): Promise<void>;
    /**
     * Check if current user has read a specific message
     */
    hasUserReadMessage(messageId: number): Promise<boolean>;
    /**
     * Get current user information for display
     */
    getCurrentUser(): IEnhancedUser | null;
    /**
     * Create a new message (admin function)
     */
    createMessage(message: Partial<IMessage>): Promise<number>;
    /**
     * Get a specific message by ID
     */
    getMessageById(messageId: number): Promise<IMessage>;
    /**
     * Check if user has access to specific functionality based on Graph groups
     */
    hasUserRole(role: string): boolean;
    private getMessagesWithProgressiveQuerying;
    private testBasicQuery;
    private testEnhancedQuery;
    private buildConservativeColumnQuery;
    private getAvailableColumns;
    private mapToMessageBasic;
    private mapToMessageWithAvailableColumns;
    private mapToMessage;
    private getRequestDigest;
    private updateMessageReadBy;
    private getDeviceInfo;
    private getMockMessages;
}
export declare const enhancedDataService: EnhancedDataService;
//# sourceMappingURL=EnhancedDataService.d.ts.map