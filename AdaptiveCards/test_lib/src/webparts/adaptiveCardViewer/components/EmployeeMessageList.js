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
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
import * as React from 'react';
import { useState, useEffect } from 'react';
import { DetailsList, SelectionMode, PrimaryButton, DefaultButton, SearchBox, Dropdown, MessageBar, MessageBarType, Panel, PanelType, Spinner, SpinnerSize } from '@fluentui/react';
import { enhancedDataService } from '../../../services/EnhancedDataService';
export var EmployeeMessageList = function (props) {
    var _a = useState([]), messages = _a[0], setMessages = _a[1];
    var _b = useState([]), filteredMessages = _b[0], setFilteredMessages = _b[1];
    var _c = useState(true), loading = _c[0], setLoading = _c[1];
    var _d = useState(null), selectedMessage = _d[0], setSelectedMessage = _d[1];
    var _e = useState(false), isPanelOpen = _e[0], setIsPanelOpen = _e[1];
    var _f = useState(''), searchText = _f[0], setSearchText = _f[1];
    var _g = useState('All'), priorityFilter = _g[0], setPriorityFilter = _g[1];
    var _h = useState(null), result = _h[0], setResult = _h[1];
    var priorityOptions = [
        { key: 'All', text: 'All Priorities' },
        { key: 'High', text: '🚨 High Priority' },
        { key: 'Medium', text: '⚠️ Medium Priority' },
        { key: 'Low', text: 'ℹ️ Low Priority' }
    ];
    useEffect(function () {
        loadMessages();
    }, []);
    useEffect(function () {
        applyFilters();
    }, [messages, searchText, priorityFilter]);
    var loadMessages = function () { return __awaiter(void 0, void 0, void 0, function () {
        var userMessages, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setLoading(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, enhancedDataService.getMessagesForCurrentUser()];
                case 2:
                    userMessages = _a.sent();
                    setMessages(userMessages);
                    setResult(null);
                    return [3 /*break*/, 5];
                case 3:
                    error_1 = _a.sent();
                    console.error('Error loading messages:', error_1);
                    setResult({
                        type: 'error',
                        message: "\u274C Failed to load messages: ".concat(error_1.message || 'Unknown error')
                    });
                    return [3 /*break*/, 5];
                case 4:
                    setLoading(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var applyFilters = function () {
        var filtered = __spreadArray([], messages, true);
        // Apply search filter
        if (searchText.trim()) {
            var searchLower_1 = searchText.toLowerCase();
            filtered = filtered.filter(function (msg) {
                return msg.Title.toLowerCase().includes(searchLower_1) ||
                    msg.MessageContent.toLowerCase().includes(searchLower_1);
            });
        }
        // Apply priority filter
        if (priorityFilter !== 'All') {
            filtered = filtered.filter(function (msg) { return msg.Priority === priorityFilter; });
        }
        // Sort by priority and creation date
        filtered.sort(function (a, b) {
            var priorityOrder = { 'High': 3, 'Medium': 2, 'Low': 1 };
            var priorityDiff = priorityOrder[b.Priority] - priorityOrder[a.Priority];
            if (priorityDiff !== 0)
                return priorityDiff;
            return new Date(b.Created).getTime() - new Date(a.Created).getTime();
        });
        setFilteredMessages(filtered);
    };
    var handleMarkAsRead = function (message) { return __awaiter(void 0, void 0, void 0, function () {
        var error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, enhancedDataService.markMessageAsRead(message.Id)];
                case 1:
                    _a.sent();
                    setResult({
                        type: 'success',
                        message: "\u2705 Message \"".concat(message.Title, "\" marked as read")
                    });
                    // Refresh messages to update read status
                    loadMessages();
                    return [3 /*break*/, 3];
                case 2:
                    error_2 = _a.sent();
                    setResult({
                        type: 'error',
                        message: "\u274C Failed to mark as read: ".concat(error_2.message)
                    });
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var handleViewMessage = function (message) {
        setSelectedMessage(message);
        setIsPanelOpen(true);
    };
    var checkIfRead = function (messageId) { return __awaiter(void 0, void 0, void 0, function () {
        var error_3;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, enhancedDataService.hasUserReadMessage(messageId)];
                case 1: return [2 /*return*/, _a.sent()];
                case 2:
                    error_3 = _a.sent();
                    return [2 /*return*/, false];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var columns = [
        {
            key: 'priority',
            name: 'Priority',
            fieldName: 'Priority',
            minWidth: 80,
            maxWidth: 80,
            onRender: function (item) {
                var priorityIcons = {
                    'High': '🚨',
                    'Medium': '⚠️',
                    'Low': 'ℹ️'
                };
                return "".concat(priorityIcons[item.Priority], " ").concat(item.Priority);
            }
        },
        {
            key: 'title',
            name: 'Message Title',
            fieldName: 'Title',
            minWidth: 200,
            maxWidth: 300,
            isResizable: true,
            onRender: function (item) { return (React.createElement("div", null,
                React.createElement("strong", null, item.Title),
                React.createElement("div", { style: { fontSize: '12px', color: '#666' } },
                    new Date(item.Created).toLocaleDateString('sv-SE'),
                    " ",
                    new Date(item.Created).toLocaleTimeString('sv-SE')))); }
        },
        {
            key: 'content',
            name: 'Content Preview',
            fieldName: 'MessageContent',
            minWidth: 250,
            isResizable: true,
            onRender: function (item) {
                // Strip HTML and truncate
                var textContent = item.MessageContent.replace(/<[^>]*>/g, '');
                var preview = textContent.length > 100 ? textContent.substring(0, 100) + '...' : textContent;
                return React.createElement("span", null, preview);
            }
        },
        {
            key: 'targetAudience',
            name: 'Target Audience',
            fieldName: 'TargetAudience',
            minWidth: 120,
            maxWidth: 150
        },
        {
            key: 'actions',
            name: 'Actions',
            minWidth: 120,
            onRender: function (item) { return (React.createElement("div", { style: { display: 'flex', gap: '8px' } },
                React.createElement(DefaultButton, { text: "View", onClick: function () { return handleViewMessage(item); }, styles: { root: { minWidth: '50px' } } }),
                React.createElement(PrimaryButton, { text: "Mark Read", onClick: function () { return handleMarkAsRead(item); }, styles: { root: { minWidth: '70px' } } }))); }
        }
    ];
    var renderMessagePanel = function () { return (React.createElement(Panel, { isOpen: isPanelOpen, onDismiss: function () { return setIsPanelOpen(false); }, type: PanelType.medium, headerText: (selectedMessage === null || selectedMessage === void 0 ? void 0 : selectedMessage.Title) || 'Message Details', closeButtonAriaLabel: "Close" }, selectedMessage && (React.createElement("div", { style: { padding: '16px' } },
        React.createElement("div", { style: { marginBottom: '16px' } },
            React.createElement("strong", null, "Priority:"),
            " ",
            selectedMessage.Priority === 'High' ? '🚨' : selectedMessage.Priority === 'Medium' ? '⚠️' : 'ℹ️',
            " ",
            selectedMessage.Priority),
        React.createElement("div", { style: { marginBottom: '16px' } },
            React.createElement("strong", null, "Target Audience:"),
            " ",
            selectedMessage.TargetAudience),
        React.createElement("div", { style: { marginBottom: '16px' } },
            React.createElement("strong", null, "Created:"),
            " ",
            new Date(selectedMessage.Created).toLocaleString('sv-SE')),
        React.createElement("div", { style: { marginBottom: '24px' } },
            React.createElement("strong", null, "Message Content:"),
            React.createElement("div", { style: {
                    marginTop: '8px',
                    padding: '12px',
                    border: '1px solid #ddd',
                    borderRadius: '4px',
                    backgroundColor: '#f9f9f9'
                }, dangerouslySetInnerHTML: { __html: selectedMessage.MessageContent } })),
        React.createElement(PrimaryButton, { text: "Mark as Read", onClick: function () {
                handleMarkAsRead(selectedMessage);
                setIsPanelOpen(false);
            } }))))); };
    if (loading) {
        return (React.createElement("div", { style: { textAlign: 'center', padding: '40px' } },
            React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading your messages..." })));
    }
    return (React.createElement("div", { style: { padding: '20px' } },
        React.createElement("h2", null, "\uD83D\uDCE8 Your Important Messages"),
        React.createElement("p", null, "Messages targeted to you and your role"),
        result && (React.createElement(MessageBar, { messageBarType: result.type === 'success' ? MessageBarType.success :
                result.type === 'error' ? MessageBarType.error : MessageBarType.info, styles: { root: { marginBottom: '16px' } }, onDismiss: function () { return setResult(null); } }, result.message)),
        React.createElement("div", { style: { display: 'flex', gap: '16px', marginBottom: '20px', alignItems: 'flex-end' } },
            React.createElement(SearchBox, { placeholder: "Search messages...", value: searchText, onChange: function (_, newValue) { return setSearchText(newValue || ''); }, styles: { root: { width: '300px' } } }),
            React.createElement(Dropdown, { placeholder: "Filter by priority", selectedKey: priorityFilter, options: priorityOptions, onChange: function (_, option) { return setPriorityFilter((option === null || option === void 0 ? void 0 : option.key) || 'All'); }, styles: { dropdown: { width: '150px' } } }),
            React.createElement(DefaultButton, { text: "Refresh", iconProps: { iconName: 'Refresh' }, onClick: loadMessages })),
        filteredMessages.length === 0 ? (React.createElement(MessageBar, { messageBarType: MessageBarType.info }, messages.length === 0
            ? "📭 No messages found. You're all caught up!"
            : "📭 No messages match your current filters.")) : (React.createElement(DetailsList, { items: filteredMessages, columns: columns, selectionMode: SelectionMode.none, layoutMode: 0, styles: {
                root: { border: '1px solid #ddd' }
            } })),
        renderMessagePanel(),
        React.createElement("div", { style: { marginTop: '30px', padding: '15px', backgroundColor: '#e8f4fd', borderRadius: '8px' } },
            React.createElement("h4", null, "\uD83D\uDCA1 Employee Message Center"),
            React.createElement("ul", null,
                React.createElement("li", null,
                    React.createElement("strong", null, "\uD83D\uDCCB View Messages:"),
                    " Messages targeted to your role and groups"),
                React.createElement("li", null,
                    React.createElement("strong", null, "\uD83D\uDD0D Search & Filter:"),
                    " Find specific messages quickly"),
                React.createElement("li", null,
                    React.createElement("strong", null, "\u2705 Mark as Read:"),
                    " Confirm you've seen important information"),
                React.createElement("li", null,
                    React.createElement("strong", null, "\uD83D\uDCCA Priority Levels:"),
                    " High \uD83D\uDEA8, Medium \u26A0\uFE0F, Low \u2139\uFE0F")))));
};
//# sourceMappingURL=EmployeeMessageList.js.map