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
import { useState, useEffect } from 'react';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { MessageBar, MessageBarType, PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
export var MarkReadPage = function (props) {
    var _a;
    var _b = useState('loading'), status = _b[0], setStatus = _b[1];
    var _c = useState(null), message = _c[0], setMessage = _c[1];
    var _d = useState(''), error = _d[0], setError = _d[1];
    useEffect(function () {
        markAsRead();
    }, []);
    var markAsRead = function () { return __awaiter(void 0, void 0, void 0, function () {
        var messageData, alreadyRead, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 4, , 5]);
                    setStatus('loading');
                    if (!enhancedDataService.getCurrentUser()) {
                        setError('Data service not available');
                        setStatus('error');
                        return [2 /*return*/];
                    }
                    return [4 /*yield*/, enhancedDataService.getMessageById(props.messageId)];
                case 1:
                    messageData = _a.sent();
                    setMessage(messageData);
                    return [4 /*yield*/, enhancedDataService.hasUserReadMessage(props.messageId)];
                case 2:
                    alreadyRead = _a.sent();
                    if (alreadyRead) {
                        setStatus('already-read');
                        return [2 /*return*/];
                    }
                    // Mark as read
                    return [4 /*yield*/, enhancedDataService.markMessageAsRead(props.messageId)];
                case 3:
                    // Mark as read
                    _a.sent();
                    setStatus('success');
                    return [3 /*break*/, 5];
                case 4:
                    err_1 = _a.sent();
                    console.error('Error marking message as read:', err_1);
                    setError(err_1.message || 'Failed to mark message as read');
                    setStatus('error');
                    return [3 /*break*/, 5];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var goToDashboard = function () {
        // Always use the correct subsite URL regardless of current context
        window.location.href = "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/SitePages/Dashboard.aspx";
    };
    var goToMessages = function () {
        // Always use the correct subsite URL regardless of current context
        window.location.href = "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/Lists/Important%20Messages/AllItems.aspx";
    };
    var goToTeams = function () {
        // If coming from Teams, try to go back
        if (props.source === 'teams') {
            window.close(); // This might work in Teams context
        }
    };
    return (React.createElement("div", { style: { padding: '20px', maxWidth: '600px', margin: '0 auto' } },
        React.createElement("h2", null, "\uD83D\uDCE8 Message Action"),
        message && (React.createElement("div", { style: {
                background: '#f8f9fa',
                padding: '15px',
                borderRadius: '8px',
                marginBottom: '20px',
                border: '1px solid #dee2e6'
            } },
            React.createElement("h3", null, message.Title),
            React.createElement("p", null, message.MessageContent),
            React.createElement("small", null,
                "From: ", (_a = message.Author) === null || _a === void 0 ? void 0 :
                _a.Title,
                " | Priority: ",
                message.Priority))),
        status === 'loading' && (React.createElement(MessageBar, { messageBarType: MessageBarType.info }, "\uD83D\uDCE4 Processing your request...")),
        status === 'success' && (React.createElement("div", null,
            React.createElement(MessageBar, { messageBarType: MessageBarType.success }, "\u2705 Message marked as read successfully! Thank you for confirming."),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement(PrimaryButton, { text: "\uD83D\uDCCA View Dashboard", onClick: goToDashboard, style: { marginRight: '10px' } }),
                React.createElement(DefaultButton, { text: "\uD83D\uDCCB All Messages", onClick: goToMessages, style: { marginRight: '10px' } }),
                props.source === 'teams' && (React.createElement(DefaultButton, { text: "\u21A9\uFE0F Back to Teams", onClick: goToTeams }))))),
        status === 'already-read' && (React.createElement("div", null,
            React.createElement(MessageBar, { messageBarType: MessageBarType.warning }, "\u2139\uFE0F You have already marked this message as read."),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement(PrimaryButton, { text: "\uD83D\uDCCA View Dashboard", onClick: goToDashboard, style: { marginRight: '10px' } }),
                React.createElement(DefaultButton, { text: "\uD83D\uDCCB All Messages", onClick: goToMessages, style: { marginRight: '10px' } }),
                props.source === 'teams' && (React.createElement(DefaultButton, { text: "\u21A9\uFE0F Back to Teams", onClick: goToTeams }))))),
        status === 'error' && (React.createElement("div", null,
            React.createElement(MessageBar, { messageBarType: MessageBarType.error },
                "\u274C Error: ",
                error),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement(PrimaryButton, { text: "\uD83D\uDD04 Try Again", onClick: markAsRead, style: { marginRight: '10px' } }),
                React.createElement(DefaultButton, { text: "\uD83D\uDCCA View Dashboard", onClick: goToDashboard, style: { marginRight: '10px' } }),
                React.createElement(DefaultButton, { text: "\uD83D\uDCCB All Messages", onClick: goToMessages })))),
        React.createElement("div", { style: {
                marginTop: '30px',
                padding: '15px',
                background: '#e8f4fd',
                borderRadius: '8px',
                fontSize: '14px'
            } },
            React.createElement("h4", null, "\uD83D\uDCDA About Message Tracking:"),
            React.createElement("ul", null,
                React.createElement("li", null, "\u2705 Your read confirmation is logged in SharePoint"),
                React.createElement("li", null, "\uD83D\uDCCA Administrators can view read statistics on the dashboard"),
                React.createElement("li", null, "\uD83D\uDD12 Only you and authorized personnel can see your read status"),
                React.createElement("li", null, "\uD83D\uDCF1 This works from Teams, email, or SharePoint")))));
};
//# sourceMappingURL=MarkReadPage.js.map