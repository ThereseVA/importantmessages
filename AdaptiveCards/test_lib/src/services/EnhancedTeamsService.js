var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { generateTeamsCard, generateMessageCard } from '../webparts/adaptiveCardViewer/models/CardTemplates';
import { enhancedDataService } from './EnhancedDataService';
import { graphService } from './GraphService';
/**
 * Enhanced Teams Distribution Service using Microsoft Graph API
 * No external webhook URLs required - uses delegated permissions
 */
var EnhancedTeamsService = /** @class */ (function () {
    function EnhancedTeamsService() {
    }
    /**
     * Send message to Teams chat using Graph API (no admin approval needed)
     */
    EnhancedTeamsService.sendToTeamsChat = function (chatId, message) {
        return __awaiter(this, void 0, void 0, function () {
            var adaptiveCard, teamsMessage, result, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 8]);
                        console.log('📤 Sending message to Teams chat via Graph API:', chatId);
                        adaptiveCard = generateTeamsCard(message);
                        teamsMessage = {
                            body: {
                                contentType: 'html',
                                content: "<h3>".concat(message.Title, "</h3><p>").concat(message.MessageContent, "</p>")
                            },
                            attachments: [{
                                    id: "message-".concat(message.Id),
                                    contentType: 'application/vnd.microsoft.card.adaptive',
                                    content: adaptiveCard,
                                    name: "Message: ".concat(message.Title)
                                }]
                        };
                        return [4 /*yield*/, graphService.sendTeamsMessage(chatId, JSON.stringify(teamsMessage))];
                    case 1:
                        result = _a.sent();
                        if (!result) return [3 /*break*/, 3];
                        console.log('✅ Successfully sent message to Teams chat');
                        return [4 /*yield*/, this.logDistribution(message.Id, 'Teams Chat', chatId, 'Success')];
                    case 2:
                        _a.sent();
                        return [2 /*return*/, true];
                    case 3:
                        console.error('❌ Failed to send message to Teams chat');
                        return [4 /*yield*/, this.logDistribution(message.Id, 'Teams Chat', chatId, 'Failed')];
                    case 4:
                        _a.sent();
                        return [2 /*return*/, false];
                    case 5: return [3 /*break*/, 8];
                    case 6:
                        error_1 = _a.sent();
                        console.error('❌ Error sending to Teams chat:', error_1);
                        return [4 /*yield*/, this.logDistribution(message.Id, 'Teams Chat', chatId, 'Error')];
                    case 7:
                        _a.sent();
                        return [2 /*return*/, false];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get user's Teams chats for message distribution
     */
    EnhancedTeamsService.getUserTeamsChats = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                try {
                    // This would use Graph API to get user's chats
                    // For now, return empty array as this requires specific Graph permissions
                    console.log('📱 Getting user Teams chats (placeholder)');
                    return [2 /*return*/, []];
                }
                catch (error) {
                    console.error('❌ Error getting Teams chats:', error);
                    return [2 /*return*/, []];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Create a simple HTML notification instead of external webhook
     */
    EnhancedTeamsService.createNotification = function (message) {
        return __awaiter(this, void 0, void 0, function () {
            var adaptiveCard, htmlNotification;
            return __generator(this, function (_a) {
                try {
                    adaptiveCard = generateTeamsCard(message);
                    htmlNotification = "\n        <div style=\"border: 1px solid #e1e5e9; border-radius: 8px; padding: 16px; margin: 8px 0; background: #f8f9fa;\">\n          <div style=\"display: flex; align-items: center; margin-bottom: 12px;\">\n            <div style=\"width: 4px; height: 24px; background: ".concat(this.getPriorityColor(message.Priority), "; margin-right: 12px; border-radius: 2px;\"></div>\n            <h3 style=\"margin: 0; color: #333; font-size: 18px;\">").concat(message.Title, "</h3>\n          </div>\n          <p style=\"margin: 8px 0; color: #666; line-height: 1.4;\">").concat(message.MessageContent, "</p>\n          <div style=\"display: flex; justify-content: space-between; align-items: center; margin-top: 12px; padding-top: 12px; border-top: 1px solid #e1e5e9;\">\n            <span style=\"background: ").concat(this.getPriorityColor(message.Priority), "; color: white; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;\">\n              ").concat(message.Priority, " Priority\n            </span>\n            <span style=\"color: #888; font-size: 12px;\">\n              ").concat(new Date(message.Created).toLocaleDateString(), "\n            </span>\n          </div>\n        </div>\n      ");
                    return [2 /*return*/, htmlNotification];
                }
                catch (error) {
                    console.error('❌ Error creating notification:', error);
                    return [2 /*return*/, "<div style=\"color: red;\">Error creating notification for message: ".concat(message.Title, "</div>")];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Generate card JSON for manual sharing or export
     */
    EnhancedTeamsService.generateCardJson = function (message) {
        try {
            var card = generateTeamsCard(message);
            return JSON.stringify(card, null, 2);
        }
        catch (error) {
            console.error('❌ Error generating card JSON:', error);
            return JSON.stringify({ error: 'Failed to generate card' }, null, 2);
        }
    };
    /**
     * Generate email-friendly card HTML
     */
    EnhancedTeamsService.generateEmailHtml = function (message) {
        try {
            var card = generateMessageCard(message);
            // Convert adaptive card to HTML for email
            return "\n        <div style=\"font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto;\">\n          <div style=\"background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px 8px 0 0;\">\n            <h2 style=\"margin: 0; font-size: 24px;\">".concat(message.Title, "</h2>\n            <p style=\"margin: 8px 0 0 0; opacity: 0.9;\">").concat(message.Priority, " Priority Message</p>\n          </div>\n          <div style=\"background: white; padding: 20px; border: 1px solid #e1e5e9; border-radius: 0 0 8px 8px;\">\n            <p style=\"margin: 0 0 16px 0; color: #333; line-height: 1.6;\">").concat(message.MessageContent, "</p>\n            <div style=\"border-top: 1px solid #e1e5e9; padding-top: 16px; color: #666; font-size: 14px;\">\n              <p style=\"margin: 0;\"><strong>Target Audience:</strong> ").concat(message.TargetAudience, "</p>\n              <p style=\"margin: 4px 0 0 0;\"><strong>Created:</strong> ").concat(new Date(message.Created).toLocaleString(), "</p>\n            </div>\n          </div>\n        </div>\n      ");
        }
        catch (error) {
            console.error('❌ Error generating email HTML:', error);
            return "<div style=\"color: red;\">Error generating email for message: ".concat(message.Title, "</div>");
        }
    };
    /**
     * Create shareable link for message (SharePoint-based)
     */
    EnhancedTeamsService.createShareableLink = function (message) {
        return __awaiter(this, void 0, void 0, function () {
            var siteUrl, encodedTitle, shareLink;
            return __generator(this, function (_a) {
                try {
                    siteUrl = enhancedDataService.getCurrentSiteUrl();
                    encodedTitle = encodeURIComponent(message.Title);
                    shareLink = "".concat(siteUrl, "/Lists/Important%20Messages/DispForm.aspx?ID=").concat(message.Id, "&Title=").concat(encodedTitle);
                    console.log('🔗 Created shareable link:', shareLink);
                    return [2 /*return*/, shareLink];
                }
                catch (error) {
                    console.error('❌ Error creating shareable link:', error);
                    return [2 /*return*/, '#'];
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Distribute message to user's accessible channels (no external webhooks)
     */
    EnhancedTeamsService.distributeToAccessibleChannels = function (message) {
        return __awaiter(this, void 0, void 0, function () {
            var result, chats, htmlNotification, shareLink, _i, chats_1, chat, success, error_2, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        result = {
                            success: 0,
                            failed: 0,
                            details: []
                        };
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 13, , 14]);
                        console.log('📊 Starting distribution to accessible channels');
                        return [4 /*yield*/, this.getUserTeamsChats()];
                    case 2:
                        chats = _a.sent();
                        if (!(chats.length === 0)) return [3 /*break*/, 6];
                        console.log('ℹ️ No accessible Teams chats found, creating alternative notifications');
                        return [4 /*yield*/, this.createNotification(message)];
                    case 3:
                        htmlNotification = _a.sent();
                        return [4 /*yield*/, this.createShareableLink(message)];
                    case 4:
                        shareLink = _a.sent();
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
                        return [4 /*yield*/, this.logDistribution(message.Id, 'Alternative', 'HTML + Link', 'Success')];
                    case 5:
                        _a.sent();
                        return [3 /*break*/, 12];
                    case 6:
                        _i = 0, chats_1 = chats;
                        _a.label = 7;
                    case 7:
                        if (!(_i < chats_1.length)) return [3 /*break*/, 12];
                        chat = chats_1[_i];
                        _a.label = 8;
                    case 8:
                        _a.trys.push([8, 10, , 11]);
                        return [4 /*yield*/, this.sendToTeamsChat(chat.id, message)];
                    case 9:
                        success = _a.sent();
                        if (success) {
                            result.success++;
                            result.details.push({
                                target: chat.displayName,
                                status: 'success'
                            });
                        }
                        else {
                            result.failed++;
                            result.details.push({
                                target: chat.displayName,
                                status: 'failed',
                                error: 'Failed to send message'
                            });
                        }
                        return [3 /*break*/, 11];
                    case 10:
                        error_2 = _a.sent();
                        result.failed++;
                        result.details.push({
                            target: chat.displayName,
                            status: 'failed',
                            error: error_2.message
                        });
                        return [3 /*break*/, 11];
                    case 11:
                        _i++;
                        return [3 /*break*/, 7];
                    case 12:
                        console.log("\uD83D\uDCCA Distribution completed: ".concat(result.success, " success, ").concat(result.failed, " failed"));
                        return [2 /*return*/, result];
                    case 13:
                        error_3 = _a.sent();
                        console.error('❌ Error in distribution:', error_3);
                        result.failed = 1;
                        result.details.push({
                            target: 'Distribution System',
                            status: 'failed',
                            error: error_3.message
                        });
                        return [2 /*return*/, result];
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get priority color for UI display
     */
    EnhancedTeamsService.getPriorityColor = function (priority) {
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
    };
    /**
     * Log distribution attempts (no external calls)
     */
    EnhancedTeamsService.logDistribution = function (messageId, platform, target, status) {
        return __awaiter(this, void 0, void 0, function () {
            var logEntry;
            return __generator(this, function (_a) {
                try {
                    logEntry = {
                        MessageId: messageId,
                        Platform: platform,
                        Target: target,
                        Status: status,
                        Timestamp: new Date().toISOString()
                    };
                    console.log('📝 Logging distribution:', logEntry);
                    // In a real implementation, this would save to SharePoint
                    // For now, just log to console to avoid permission issues
                }
                catch (error) {
                    console.warn('⚠️ Could not log distribution:', error);
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Create a copy-pasteable Teams message
     */
    EnhancedTeamsService.createCopyPasteMessage = function (message) {
        try {
            return "\n\uD83D\uDCE2 **".concat(message.Title, "**\n\n").concat(message.MessageContent, "\n\n\uD83C\uDFAF **Target Audience:** ").concat(message.TargetAudience, "\n\u26A1 **Priority:** ").concat(message.Priority, "\n\uD83D\uDCC5 **Created:** ").concat(new Date(message.Created).toLocaleDateString(), "\n\n---\n*This message was generated by the SPFx Adaptive Cards solution*\n      ").trim();
        }
        catch (error) {
            console.error('❌ Error creating copy-paste message:', error);
            return "Error creating message: ".concat(message.Title);
        }
    };
    /**
     * Check if Teams integration is available (without requiring permissions)
     */
    EnhancedTeamsService.checkTeamsAvailability = function () {
        return __awaiter(this, void 0, void 0, function () {
            var isTeamsContext, isGraphAvailable;
            return __generator(this, function (_a) {
                try {
                    isTeamsContext = window.location.href.includes('teams.microsoft.com') ||
                        window.location.href.includes('teams.office.com');
                    isGraphAvailable = graphService.isInitialized();
                    console.log('🔍 Teams availability check:', { isTeamsContext: isTeamsContext, isGraphAvailable: isGraphAvailable });
                    return [2 /*return*/, isTeamsContext || isGraphAvailable];
                }
                catch (error) {
                    console.error('❌ Error checking Teams availability:', error);
                    return [2 /*return*/, false];
                }
                return [2 /*return*/];
            });
        });
    };
    return EnhancedTeamsService;
}());
export { EnhancedTeamsService };
// Export singleton instance if needed
//# sourceMappingURL=EnhancedTeamsService.js.map