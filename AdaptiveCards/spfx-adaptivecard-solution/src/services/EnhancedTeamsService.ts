import { generateTeamsCard, generateMessageCard } from '../webparts/adaptiveCardViewer/models/CardTemplates';
import { enhancedDataService, IMessage } from './EnhancedDataService';
import { graphService } from './GraphService';

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
export class EnhancedTeamsService {
  
  /**
   * Send message to Teams chat using Graph API (no admin approval needed)
   */
  public static async sendToTeamsChat(chatId: string, message: IMessage): Promise<boolean> {
    try {
      console.log('📤 Sending message to Teams chat via Graph API:', chatId);
      
      // Generate adaptive card content
      const adaptiveCard = generateTeamsCard(message);
      
      // Create Teams message with adaptive card attachment
      const teamsMessage = {
        body: {
          contentType: 'html' as const,
          content: `<h3>${message.Title}</h3><p>${message.MessageContent}</p>`
        },
        attachments: [{
          id: `message-${message.Id}`,
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: adaptiveCard,
          name: `Message: ${message.Title}`
        }]
      };

      // Use Graph service to send message
      const result = await graphService.sendTeamsMessage(chatId, JSON.stringify(teamsMessage));
      
      if (result) {
        console.log('✅ Successfully sent message to Teams chat');
        await this.logDistribution(message.Id, 'Teams Chat', chatId, 'Success');
        return true;
      } else {
        console.error('❌ Failed to send message to Teams chat');
        await this.logDistribution(message.Id, 'Teams Chat', chatId, 'Failed');
        return false;
      }
    } catch (error) {
      console.error('❌ Error sending to Teams chat:', error);
      await this.logDistribution(message.Id, 'Teams Chat', chatId, 'Error');
      return false;
    }
  }

  /**
   * Get user's Teams chats for message distribution
   */
  public static async getUserTeamsChats(): Promise<Array<{ id: string; displayName: string; members: string[] }>> {
    try {
      // This would use Graph API to get user's chats
      // For now, return empty array as this requires specific Graph permissions
      console.log('📱 Getting user Teams chats (placeholder)');
      return [];
    } catch (error) {
      console.error('❌ Error getting Teams chats:', error);
      return [];
    }
  }

  /**
   * Create a simple HTML notification instead of external webhook
   */
  public static async createNotification(message: IMessage): Promise<string> {
    try {
      const adaptiveCard = generateTeamsCard(message);
      
      // Create an HTML representation of the adaptive card
      const htmlNotification = `
        <div style="border: 1px solid #e1e5e9; border-radius: 8px; padding: 16px; margin: 8px 0; background: #f8f9fa;">
          <div style="display: flex; align-items: center; margin-bottom: 12px;">
            <div style="width: 4px; height: 24px; background: ${this.getPriorityColor(message.Priority)}; margin-right: 12px; border-radius: 2px;"></div>
            <h3 style="margin: 0; color: #333; font-size: 18px;">${message.Title}</h3>
          </div>
          <p style="margin: 8px 0; color: #666; line-height: 1.4;">${message.MessageContent}</p>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px; padding-top: 12px; border-top: 1px solid #e1e5e9;">
            <span style="background: ${this.getPriorityColor(message.Priority)}; color: white; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;">
              ${message.Priority} Priority
            </span>
            <span style="color: #888; font-size: 12px;">
              ${new Date(message.Created).toLocaleDateString()}
            </span>
          </div>
        </div>
      `;
      
      return htmlNotification;
    } catch (error) {
      console.error('❌ Error creating notification:', error);
      return `<div style="color: red;">Error creating notification for message: ${message.Title}</div>`;
    }
  }

  /**
   * Generate card JSON for manual sharing or export
   */
  public static generateCardJson(message: IMessage): string {
    try {
      const card = generateTeamsCard(message);
      return JSON.stringify(card, null, 2);
    } catch (error) {
      console.error('❌ Error generating card JSON:', error);
      return JSON.stringify({ error: 'Failed to generate card' }, null, 2);
    }
  }

  /**
   * Generate email-friendly card HTML
   */
  public static generateEmailHtml(message: IMessage): string {
    try {
      const card = generateMessageCard(message);
      
      // Convert adaptive card to HTML for email
      return `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px 8px 0 0;">
            <h2 style="margin: 0; font-size: 24px;">${message.Title}</h2>
            <p style="margin: 8px 0 0 0; opacity: 0.9;">${message.Priority} Priority Message</p>
          </div>
          <div style="background: white; padding: 20px; border: 1px solid #e1e5e9; border-radius: 0 0 8px 8px;">
            <p style="margin: 0 0 16px 0; color: #333; line-height: 1.6;">${message.MessageContent}</p>
            <div style="border-top: 1px solid #e1e5e9; padding-top: 16px; color: #666; font-size: 14px;">
              <p style="margin: 0;"><strong>Target Audience:</strong> ${message.TargetAudience}</p>
              <p style="margin: 4px 0 0 0;"><strong>Created:</strong> ${new Date(message.Created).toLocaleString()}</p>
            </div>
          </div>
        </div>
      `;
    } catch (error) {
      console.error('❌ Error generating email HTML:', error);
      return `<div style="color: red;">Error generating email for message: ${message.Title}</div>`;
    }
  }

  /**
   * Create shareable link for message (SharePoint-based)
   */
  public static async createShareableLink(message: IMessage): Promise<string> {
    try {
      const siteUrl = enhancedDataService.getCurrentSiteUrl();
      const encodedTitle = encodeURIComponent(message.Title);
      
      // Create a link to the SharePoint list item or a custom page
      const shareLink = `${siteUrl}/Lists/Important%20Messages/DispForm.aspx?ID=${message.Id}&Title=${encodedTitle}`;
      
      console.log('🔗 Created shareable link:', shareLink);
      return shareLink;
    } catch (error) {
      console.error('❌ Error creating shareable link:', error);
      return '#';
    }
  }

  /**
   * Distribute message to user's accessible channels (no external webhooks)
   */
  public static async distributeToAccessibleChannels(message: IMessage): Promise<IDistributionResult> {
    const result: IDistributionResult = {
      success: 0,
      failed: 0,
      details: []
    };

    try {
      console.log('📊 Starting distribution to accessible channels');
      
      // Get user's accessible Teams chats
      const chats = await this.getUserTeamsChats();
      
      if (chats.length === 0) {
        console.log('ℹ️ No accessible Teams chats found, creating alternative notifications');
        
        // Create alternative distribution methods
        const htmlNotification = await this.createNotification(message);
        const shareLink = await this.createShareableLink(message);
        
        result.details.push({
          target: 'HTML Notification',
          status: 'success'
        });
        
        result.details.push({
          target: 'Shareable Link',
          status: 'success'
        });
        
        result.success = 2;
        
        console.log('✅ Created alternative distribution methods');
        await this.logDistribution(message.Id, 'Alternative', 'HTML + Link', 'Success');
      } else {
        // Distribute to available chats
        for (const chat of chats) {
          try {
            const success = await this.sendToTeamsChat(chat.id, message);
            if (success) {
              result.success++;
              result.details.push({
                target: chat.displayName,
                status: 'success'
              });
            } else {
              result.failed++;
              result.details.push({
                target: chat.displayName,
                status: 'failed',
                error: 'Failed to send message'
              });
            }
          } catch (error) {
            result.failed++;
            result.details.push({
              target: chat.displayName,
              status: 'failed',
              error: error.message
            });
          }
        }
      }
      
      console.log(`📊 Distribution completed: ${result.success} success, ${result.failed} failed`);
      return result;
      
    } catch (error) {
      console.error('❌ Error in distribution:', error);
      result.failed = 1;
      result.details.push({
        target: 'Distribution System',
        status: 'failed',
        error: error.message
      });
      return result;
    }
  }

  /**
   * Get priority color for UI display
   */
  private static getPriorityColor(priority: string): string {
    switch (priority.toLowerCase()) {
      case 'high':
        return '#d73502';
      case 'medium':
        return '#f7630c';
      case 'low':
        return '#0f7b0f';
      default:
        return '#666';
    }
  }

  /**
   * Log distribution attempts (no external calls)
   */
  private static async logDistribution(messageId: number, platform: string, target: string, status: string): Promise<void> {
    try {
      // Use the enhanced data service which doesn't require external permissions
      const logEntry = {
        MessageId: messageId,
        Platform: platform,
        Target: target,
        Status: status,
        Timestamp: new Date().toISOString()
      };
      
      console.log('📝 Logging distribution:', logEntry);
      
      // In a real implementation, this would save to SharePoint
      // For now, just log to console to avoid permission issues
    } catch (error) {
      console.warn('⚠️ Could not log distribution:', error);
    }
  }

  /**
   * Create a copy-pasteable Teams message
   */
  public static createCopyPasteMessage(message: IMessage): string {
    try {
      return `
📢 **${message.Title}**

${message.MessageContent}

🎯 **Target Audience:** ${message.TargetAudience}
⚡ **Priority:** ${message.Priority}
📅 **Created:** ${new Date(message.Created).toLocaleDateString()}

---
*This message was generated by the SPFx Adaptive Cards solution*
      `.trim();
    } catch (error) {
      console.error('❌ Error creating copy-paste message:', error);
      return `Error creating message: ${message.Title}`;
    }
  }

  /**
   * Check if Teams integration is available (without requiring permissions)
   */
  public static async checkTeamsAvailability(): Promise<boolean> {
    try {
      // Check if we're in Teams context or if Graph service is available
      const isTeamsContext = window.location.href.includes('teams.microsoft.com') || 
                           window.location.href.includes('teams.office.com');
      
      const isGraphAvailable = graphService.isInitialized();
      
      console.log('🔍 Teams availability check:', { isTeamsContext, isGraphAvailable });
      
      return isTeamsContext || isGraphAvailable;
    } catch (error) {
      console.error('❌ Error checking Teams availability:', error);
      return false;
    }
  }
}

// Export singleton instance if needed
