import { IMessage } from './EnhancedDataService';
export interface ITeamsChannelMessage {
    body: {
        contentType: 'html' | 'text';
        content: string;
    };
    attachments?: Array<{
        id: string;
        contentType: string;
        content: any;
        name?: string;
    }>;
}
export interface IDistributionResult {
    success: number;
    failed: number;
    details: Array<{
        target: string;
        status: 'success' | 'failed';
        error?: string;
    }>;
}
/**
 * Enhanced Teams Distribution Service using Microsoft Graph API
 * No external webhook URLs required - uses delegated permissions
 */
export declare class EnhancedTeamsService {
    /**
     * Send message to Teams chat using Graph API (no admin approval needed)
     */
    static sendToTeamsChat(chatId: string, message: IMessage): Promise<boolean>;
    /**
     * Get user's Teams chats for message distribution
     */
    static getUserTeamsChats(): Promise<Array<{
        id: string;
        displayName: string;
        members: string[];
    }>>;
    /**
     * Create a simple HTML notification instead of external webhook
     */
    static createNotification(message: IMessage): Promise<string>;
    /**
     * Generate card JSON for manual sharing or export
     */
    static generateCardJson(message: IMessage): string;
    /**
     * Generate email-friendly card HTML
     */
    static generateEmailHtml(message: IMessage): string;
    /**
     * Create shareable link for message (SharePoint-based)
     */
    static createShareableLink(message: IMessage): Promise<string>;
    /**
     * Distribute message to user's accessible channels (no external webhooks)
     */
    static distributeToAccessibleChannels(message: IMessage): Promise<IDistributionResult>;
    /**
     * Get priority color for UI display
     */
    private static getPriorityColor;
    /**
     * Log distribution attempts (no external calls)
     */
    private static logDistribution;
    /**
     * Create a copy-pasteable Teams message
     */
    static createCopyPasteMessage(message: IMessage): string;
    /**
     * Check if Teams integration is available (without requiring permissions)
     */
    static checkTeamsAvailability(): Promise<boolean>;
}
//# sourceMappingURL=EnhancedTeamsService.d.ts.map