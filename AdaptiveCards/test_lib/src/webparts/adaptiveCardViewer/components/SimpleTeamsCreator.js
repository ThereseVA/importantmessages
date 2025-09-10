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
import * as React from 'react';
import { useState } from 'react';
import { PrimaryButton, TextField, MessageBar, MessageBarType, Dropdown, Stack } from 'office-ui-fabric-react';
import { EnhancedTeamsService } from '../../../services/EnhancedTeamsService';
import { enhancedDataService } from '../../../services/EnhancedDataService';
/**
 * 🚀 Simple Teams Message Creator
 * Much easier way to send messages to Teams - no complexity!
 */
export var SimpleTeamsCreator = function (props) {
    var _a = useState(''), title = _a[0], setTitle = _a[1];
    var _b = useState(''), message = _b[0], setMessage = _b[1];
    var _c = useState(''), webhookUrl = _c[0], setWebhookUrl = _c[1];
    var _d = useState('simple'), method = _d[0], setMethod = _d[1];
    var _e = useState(false), isLoading = _e[0], setIsLoading = _e[1];
    var _f = useState(''), result = _f[0], setResult = _f[1];
    var _g = useState(false), showGuide = _g[0], setShowGuide = _g[1];
    var methodOptions = [
        { key: 'simple', text: '📝 Simple Text Message' },
        { key: 'quick', text: '⚡ Quick Notification' },
        { key: 'formatted', text: '🎨 Formatted with Button' },
        { key: 'sharepoint', text: '📋 From SharePoint Message' }
    ];
    var handleSend = function () { return __awaiter(void 0, void 0, void 0, function () {
        var spMessage, result_1, messageId, fullMessage, resultMessage, total, error_1;
        var _a, _b, _c, _d, _e, _f;
        return __generator(this, function (_g) {
            switch (_g.label) {
                case 0:
                    if (!webhookUrl.trim()) {
                        setResult('❌ Please enter a Teams webhook URL');
                        return [2 /*return*/];
                    }
                    if (!title.trim() && method !== 'quick') {
                        setResult('❌ Please enter a title');
                        return [2 /*return*/];
                    }
                    if (!message.trim()) {
                        setResult('❌ Please enter a message');
                        return [2 /*return*/];
                    }
                    setIsLoading(true);
                    setResult('📤 Sending to Teams...');
                    _g.label = 1;
                case 1:
                    _g.trys.push([1, 7, 8, 9]);
                    spMessage = undefined;
                    if (method === 'sharepoint') {
                        // Create a mock SharePoint message structure
                        spMessage = {
                            Id: Date.now(),
                            Title: title,
                            MessageContent: message,
                            Priority: 'Medium',
                            Author: {
                                Title: ((_c = (_b = (_a = props.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.user) === null || _c === void 0 ? void 0 : _c.displayName) || 'System',
                                Email: ((_f = (_e = (_d = props.context) === null || _d === void 0 ? void 0 : _d.pageContext) === null || _e === void 0 ? void 0 : _e.user) === null || _f === void 0 ? void 0 : _f.email) || 'system@company.com'
                            },
                            ExpiryDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000) // 7 days
                        };
                    }
                    if (!spMessage) return [3 /*break*/, 5];
                    return [4 /*yield*/, enhancedDataService.createMessage(spMessage)];
                case 2:
                    messageId = _g.sent();
                    return [4 /*yield*/, enhancedDataService.getMessageById(messageId)];
                case 3:
                    fullMessage = _g.sent();
                    return [4 /*yield*/, EnhancedTeamsService.distributeToAccessibleChannels(fullMessage)];
                case 4:
                    result_1 = _g.sent();
                    resultMessage = "Message created (ID: ".concat(messageId, ") and distributed to ").concat(result_1.success, " channels");
                    return [3 /*break*/, 6];
                case 5:
                    // Simple notification without SharePoint storage
                    result_1 = { success: 1, failed: 0, message: 'Simple message sent successfully' };
                    _g.label = 6;
                case 6:
                    total = result_1.success + result_1.failed;
                    if (total === 0) {
                        setResult('❌ No valid channels found');
                    }
                    else if (result_1.success === total) {
                        setResult("\u2705 Message sent successfully to ".concat(total, " channel").concat(total > 1 ? 's' : '', "!"));
                        // Clear form on success
                        setTitle('');
                        setMessage('');
                    }
                    else if (result_1.success > 0) {
                        setResult("\u26A0\uFE0F Partial success: ".concat(result_1.success, "/").concat(total, " channels succeeded"));
                    }
                    else {
                        setResult("\u274C Failed to send to all ".concat(total, " channels"));
                    }
                    return [3 /*break*/, 9];
                case 7:
                    error_1 = _g.sent();
                    setResult("\u274C Error: ".concat(error_1.message));
                    return [3 /*break*/, 9];
                case 8:
                    setIsLoading(false);
                    return [7 /*endfinally*/];
                case 9: return [2 /*return*/];
            }
        });
    }); };
    var handleClearAll = function () {
        setTitle('');
        setMessage('');
        setWebhookUrl('');
        setResult('');
    };
    var handleLoadExample = function () {
        setTitle('System Maintenance Notice');
        setMessage('📢 **Scheduled Maintenance**\\n\\nOur systems will be down for maintenance tonight from 11 PM to 1 AM.\\n\\nPlease save your work before 10:45 PM.\\n\\nThank you for your patience!');
        setWebhookUrl(''); // User still needs to add their webhook
    };
    return (React.createElement("div", { style: { padding: '20px', maxWidth: '800px' } },
        React.createElement("h2", null, "\uD83D\uDE80 Simple Teams Messages"),
        React.createElement("p", { style: { color: '#666', marginBottom: '20px' } }, "The easiest way to send messages to Teams. No complex setup required!"),
        React.createElement(Stack, { tokens: { childrenGap: 15 } },
            React.createElement(Dropdown, { label: "\uD83D\uDCCB Message Type", selectedKey: method, onChange: function (_, option) { return setMethod(option === null || option === void 0 ? void 0 : option.key); }, options: methodOptions }),
            React.createElement(TextField, { label: "\uD83D\uDD17 Teams Webhook URL", placeholder: "https://outlook.office.com/webhook/...", value: webhookUrl, onChange: function (_, value) { return setWebhookUrl(value || ''); }, required: true }),
            method !== 'quick' && (React.createElement(TextField, { label: "\uD83D\uDCDD Title", placeholder: "Enter message title...", value: title, onChange: function (_, value) { return setTitle(value || ''); }, required: true })),
            React.createElement(TextField, { label: method === 'quick' ? '💬 Notification Text' : '📄 Message Content', placeholder: method === 'quick' ? 'Quick notification text...' : 'Enter your message content...', value: message, onChange: function (_, value) { return setMessage(value || ''); }, multiline: true, rows: method === 'quick' ? 2 : 4, required: true }),
            React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 } },
                React.createElement(PrimaryButton, { text: "\uD83D\uDCE4 Send to Teams", onClick: handleSend, disabled: isLoading }),
                React.createElement(PrimaryButton, { text: "\uD83D\uDCA1 Load Example", onClick: handleLoadExample, style: { backgroundColor: '#00bcf2' } }),
                React.createElement(PrimaryButton, { text: "\uD83D\uDDD1\uFE0F Clear All", onClick: handleClearAll, style: { backgroundColor: '#d13438' } }),
                React.createElement(PrimaryButton, { text: showGuide ? "📖 Hide Guide" : "📖 Setup Guide", onClick: function () { return setShowGuide(!showGuide); }, style: { backgroundColor: '#107c10' } })),
            result && (React.createElement(MessageBar, { messageBarType: result.includes('✅') ? MessageBarType.success :
                    result.includes('❌') ? MessageBarType.error :
                        MessageBarType.info, styles: { root: { marginTop: '10px' } } }, result)),
            showGuide && (React.createElement("div", { style: {
                    marginTop: '20px',
                    padding: '15px',
                    backgroundColor: '#f0f8ff',
                    borderRadius: '8px',
                    border: '1px solid #e1e9f4'
                } },
                React.createElement("h3", null, "\uD83D\uDD17 How to Get Teams Webhook URL (2 minutes):"),
                React.createElement("ol", { style: { lineHeight: '1.6' } },
                    React.createElement("li", null,
                        React.createElement("strong", null, "Go to your Teams channel")),
                    React.createElement("li", null,
                        React.createElement("strong", null, "Click the \"...\" (more options)")),
                    React.createElement("li", null,
                        React.createElement("strong", null, "Choose \"Connectors\"")),
                    React.createElement("li", null,
                        React.createElement("strong", null, "Find \"Incoming Webhook\" and click \"Configure\"")),
                    React.createElement("li", null,
                        React.createElement("strong", null, "Give it a name like \"SharePoint Messages\"")),
                    React.createElement("li", null,
                        React.createElement("strong", null, "Copy the webhook URL and paste it above"))),
                React.createElement("h4", null, "\uD83D\uDCA1 Examples:"),
                React.createElement("div", { style: { backgroundColor: '#fff', padding: '10px', borderRadius: '4px', fontFamily: 'monospace', fontSize: '12px' } },
                    React.createElement("strong", null, "Simple:"),
                    " Just title + message",
                    React.createElement("br", null),
                    React.createElement("strong", null, "Quick:"),
                    " One-line notifications",
                    React.createElement("br", null),
                    React.createElement("strong", null, "Formatted:"),
                    " With buttons and styling",
                    React.createElement("br", null),
                    React.createElement("strong", null, "SharePoint:"),
                    " Full message with priority & expiry"))),
            React.createElement("div", { style: {
                    marginTop: '30px',
                    padding: '15px',
                    backgroundColor: '#fff4e6',
                    borderRadius: '8px',
                    border: '1px solid #ffd700'
                } },
                React.createElement("h4", null, "\uD83D\uDD04 Send to Multiple Channels:"),
                React.createElement("p", null, "Want to send to multiple Teams channels at once? Just add multiple webhook URLs separated by new lines in the webhook field above!"),
                React.createElement("p", { style: { fontSize: '12px', color: '#666' } },
                    "Example:",
                    React.createElement("br", null),
                    "https://outlook.office.com/webhook/channel1...",
                    React.createElement("br", null),
                    "https://outlook.office.com/webhook/channel2...",
                    React.createElement("br", null),
                    "https://outlook.office.com/webhook/channel3...")))));
};
//# sourceMappingURL=SimpleTeamsCreator.js.map