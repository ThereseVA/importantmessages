var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { DetailsList, DetailsListLayoutMode, SelectionMode, MessageBar, MessageBarType, PrimaryButton, DefaultButton, SearchBox, Dropdown, ProgressIndicator, Panel, PanelType } from 'office-ui-fabric-react';
export var ManagerDashboard = function (props) {
    var _a = useState([]), messages = _a[0], setMessages = _a[1];
    var _b = useState([]), filteredMessages = _b[0], setFilteredMessages = _b[1];
    var _c = useState(true), loading = _c[0], setLoading = _c[1];
    var _d = useState(null), error = _d[0], setError = _d[1];
    var _e = useState(null), selectedMessage = _e[0], setSelectedMessage = _e[1];
    var _f = useState(false), isPanelOpen = _f[0], setIsPanelOpen = _f[1];
    var _g = useState(''), searchText = _g[0], setSearchText = _g[1];
    var _h = useState('All'), statusFilter = _h[0], setStatusFilter = _h[1];
    var _j = useState('All'), sourceFilter = _j[0], setSourceFilter = _j[1];
    console.log('🔧 ManagerDashboard: Component rendered, enhanced data service available:', !!enhancedDataService);
    var statusOptions = [
        { key: 'All', text: 'All Messages' },
        { key: 'Not Started', text: '🔴 Not Started (0% read)' },
        { key: 'In Progress', text: '🟡 In Progress (1-99% read)' },
        { key: 'Completed', text: '🟢 Completed (100% read)' }
    ];
    var sourceOptions = [
        { key: 'All', text: 'All Sources' },
        { key: 'Teams', text: '👥 Teams Messages' },
        { key: 'SharePoint', text: '📋 SharePoint Messages' },
        { key: 'Outlook', text: '📧 Outlook Emails' }
    ];
    useEffect(function () {
        if (!enhancedDataService.getCurrentUser()) {
            console.error('🔧 ManagerDashboard: Enhanced data service not initialized');
            setError('Enhanced data service not initialized');
            setLoading(false);
            return;
        }
        loadMessages();
    }, []);
    useEffect(function () {
        applyFilters();
    }, [messages, searchText, statusFilter, sourceFilter]);
    var loadMessages = function () { return __awaiter(void 0, void 0, void 0, function () {
        var allMessages, messagesWithStats, error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setLoading(true);
                    setError(null);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 4, 5, 6]);
                    console.log('🔧 ManagerDashboard: Loading all messages for manager view');
                    if (!enhancedDataService.getCurrentUser()) {
                        throw new Error('Enhanced data service not available');
                    }
                    return [4 /*yield*/, enhancedDataService.getActiveMessages()];
                case 2:
                    allMessages = _a.sent();
                    console.log('🔧 ManagerDashboard: Retrieved', allMessages.length, 'messages');
                    return [4 /*yield*/, Promise.all(allMessages.map(function (message) { return __awaiter(void 0, void 0, void 0, function () {
                            var stats_1, estimatedTargetUsers, readPercentage, readStatus;
                            return __generator(this, function (_a) {
                                try {
                                    stats_1 = {
                                        totalReads: 0,
                                        uniqueReaders: 0,
                                        readPercentage: 0,
                                        unreadCount: 0,
                                        readActions: []
                                    };
                                    estimatedTargetUsers = message.TargetAudience === 'All Users' ? 100 :
                                        message.TargetAudience.includes('Department') ? 25 : 50;
                                    readPercentage = Math.min((stats_1.uniqueReaders / estimatedTargetUsers) * 100, 100);
                                    readStatus = void 0;
                                    if (readPercentage === 0)
                                        readStatus = 'Not Started';
                                    else if (readPercentage < 100)
                                        readStatus = 'In Progress';
                                    else
                                        readStatus = 'Completed';
                                    return [2 /*return*/, __assign(__assign({}, message), { totalReads: stats_1.totalReads, uniqueReaders: stats_1.uniqueReaders, readPercentage: Math.round(readPercentage), lastReadDate: stats_1.readActions.length > 0 ? stats_1.readActions[0].ReadTimestamp : undefined, readStatus: readStatus, notReadUsers: [] // You could calculate this based on your user directory
                                         })];
                                }
                                catch (error) {
                                    console.error('🔧 ManagerDashboard: Error getting stats for message', message.Id, error);
                                    // Return message with default stats if stats retrieval fails
                                    return [2 /*return*/, __assign(__assign({}, message), { totalReads: 0, uniqueReaders: 0, readPercentage: 0, readStatus: 'Not Started', notReadUsers: [] })];
                                }
                                return [2 /*return*/];
                            });
                        }); }))];
                case 3:
                    messagesWithStats = _a.sent();
                    console.log('🔧 ManagerDashboard: Processed', messagesWithStats.length, 'messages with stats');
                    setMessages(messagesWithStats);
                    return [3 /*break*/, 6];
                case 4:
                    error_1 = _a.sent();
                    console.error('🔧 ManagerDashboard: Error loading messages with stats:', error_1);
                    setError("Failed to load messages: ".concat(error_1.message || 'Unknown error'));
                    // Set empty array instead of leaving in loading state
                    setMessages([]);
                    return [3 /*break*/, 6];
                case 5:
                    setLoading(false);
                    return [7 /*endfinally*/];
                case 6: return [2 /*return*/];
            }
        });
    }); };
    var applyFilters = function () {
        var filtered = __spreadArray([], messages, true);
        // Apply search filter
        if (searchText) {
            filtered = filtered.filter(function (msg) {
                return msg.Title.toLowerCase().includes(searchText.toLowerCase()) ||
                    (msg.MessageContent && msg.MessageContent.toLowerCase().includes(searchText.toLowerCase()));
            });
        }
        // Apply status filter
        if (statusFilter !== 'All') {
            filtered = filtered.filter(function (msg) { return msg.readStatus === statusFilter; });
        }
        // Apply source filter
        if (sourceFilter !== 'All') {
            filtered = filtered.filter(function (msg) { return (msg.Source || 'SharePoint') === sourceFilter; });
        }
        setFilteredMessages(filtered);
    };
    var columns = [
        {
            key: 'status',
            name: 'Status',
            fieldName: 'readStatus',
            minWidth: 80,
            maxWidth: 120,
            onRender: function (item) {
                var icon = item.readStatus === 'Completed' ? '🟢' :
                    item.readStatus === 'In Progress' ? '🟡' : '🔴';
                return React.createElement("span", null,
                    icon,
                    " ",
                    item.readStatus);
            }
        },
        {
            key: 'source',
            name: 'Source',
            fieldName: 'Source',
            minWidth: 80,
            maxWidth: 100,
            onRender: function (item) {
                var source = item.Source || 'SharePoint';
                var icon = source === 'Teams' ? '👥' : source === 'Outlook' ? '📧' : '📋';
                return React.createElement("span", null,
                    icon,
                    " ",
                    source);
            }
        },
        {
            key: 'priority',
            name: 'Priority',
            fieldName: 'Priority',
            minWidth: 80,
            maxWidth: 100,
            onRender: function (item) {
                var icon = item.Priority === 'High' ? '🚨' :
                    item.Priority === 'Medium' ? '⚠️' : 'ℹ️';
                return React.createElement("span", null,
                    icon,
                    " ",
                    item.Priority);
            }
        },
        {
            key: 'title',
            name: 'Message Title',
            fieldName: 'Title',
            minWidth: 200,
            maxWidth: 400,
            isResizable: true,
            onRender: function (item) { return (React.createElement("div", null,
                React.createElement("strong", null, item.Title),
                React.createElement("div", { style: { fontSize: '12px', color: '#666' } }, item.MessageContent ? item.MessageContent.substring(0, 100) + '...' : 'No content'))); }
        },
        {
            key: 'readProgress',
            name: 'Read Progress',
            minWidth: 150,
            maxWidth: 200,
            onRender: function (item) { return (React.createElement("div", null,
                React.createElement(ProgressIndicator, { percentComplete: item.readPercentage / 100, description: "".concat(item.readPercentage, "% (").concat(item.uniqueReaders, " users)") }))); }
        },
        {
            key: 'created',
            name: 'Created',
            fieldName: 'Created',
            minWidth: 100,
            maxWidth: 150,
            onRender: function (item) { return item.Created.toLocaleDateString(); }
        },
        {
            key: 'target',
            name: 'Target Audience',
            fieldName: 'TargetAudience',
            minWidth: 120,
            maxWidth: 180
        },
        {
            key: 'actions',
            name: 'Actions',
            minWidth: 100,
            maxWidth: 150,
            onRender: function (item) { return (React.createElement(DefaultButton, { text: "\uD83D\uDC65 View Details", onClick: function () {
                    setSelectedMessage(item);
                    setIsPanelOpen(true);
                } })); }
        }
    ];
    var getOverallStats = function () {
        var total = messages.length;
        var completed = messages.filter(function (m) { return m.readStatus === 'Completed'; }).length;
        var inProgress = messages.filter(function (m) { return m.readStatus === 'In Progress'; }).length;
        var notStarted = messages.filter(function (m) { return m.readStatus === 'Not Started'; }).length;
        var avgReadRate = total > 0 ? Math.round(messages.reduce(function (sum, m) { return sum + m.readPercentage; }, 0) / total) : 0;
        return { total: total, completed: completed, inProgress: inProgress, notStarted: notStarted, avgReadRate: avgReadRate };
    };
    var stats = getOverallStats();
    if (error) {
        return (React.createElement("div", { style: { padding: '20px' } },
            React.createElement("h2", null, "\uD83D\uDCCA Manager Dashboard - Error"),
            React.createElement(MessageBar, { messageBarType: MessageBarType.error }, error),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement(DefaultButton, { text: "Retry", onClick: loadMessages, iconProps: { iconName: 'Refresh' } }))));
    }
    if (loading) {
        return (React.createElement("div", { style: { padding: '20px' } },
            React.createElement("h2", null, "\uD83D\uDCCA Manager Dashboard - Loading..."),
            React.createElement(ProgressIndicator, { description: "Loading message statistics..." })));
    }
    return (React.createElement("div", { style: { padding: '20px' } },
        React.createElement("div", { style: { marginBottom: '20px' } },
            React.createElement("h2", { style: { color: '#323130', marginBottom: '10px' } }, "\uD83D\uDCCA Unified Message Dashboard"),
            React.createElement("div", { style: { color: '#605e5c', fontSize: '14px' } }, "View all messages from Teams and SharePoint with read tracking analytics")),
        React.createElement("div", { style: {
                display: 'grid',
                gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
                gap: '15px',
                marginBottom: '20px'
            } },
            React.createElement("div", { style: { background: '#f0f8ff', padding: '15px', borderRadius: '8px', textAlign: 'center' } },
                React.createElement("h3", { style: { margin: '0 0 5px 0', color: '#0078d4' } }, stats.total),
                React.createElement("div", null, "Total Messages")),
            React.createElement("div", { style: { background: '#f0fff0', padding: '15px', borderRadius: '8px', textAlign: 'center' } },
                React.createElement("h3", { style: { margin: '0 0 5px 0', color: '#107c10' } },
                    "\uD83D\uDFE2 ",
                    stats.completed),
                React.createElement("div", null, "Fully Read")),
            React.createElement("div", { style: { background: '#fffbf0', padding: '15px', borderRadius: '8px', textAlign: 'center' } },
                React.createElement("h3", { style: { margin: '0 0 5px 0', color: '#f7630c' } },
                    "\uD83D\uDFE1 ",
                    stats.inProgress),
                React.createElement("div", null, "Partially Read")),
            React.createElement("div", { style: { background: '#fff0f0', padding: '15px', borderRadius: '8px', textAlign: 'center' } },
                React.createElement("h3", { style: { margin: '0 0 5px 0', color: '#d13438' } },
                    "\uD83D\uDD34 ",
                    stats.notStarted),
                React.createElement("div", null, "Not Started")),
            React.createElement("div", { style: { background: '#f8f9fa', padding: '15px', borderRadius: '8px', textAlign: 'center' } },
                React.createElement("h3", { style: { margin: '0 0 5px 0', color: '#323130' } },
                    stats.avgReadRate,
                    "%"),
                React.createElement("div", null, "Avg Read Rate"))),
        React.createElement("div", { style: { display: 'flex', gap: '15px', marginBottom: '20px', alignItems: 'end' } },
            React.createElement(SearchBox, { placeholder: "Search messages...", value: searchText, onChange: function (_, newValue) { return setSearchText(newValue || ''); }, styles: { root: { width: '300px' } } }),
            React.createElement(Dropdown, { label: "Filter by Status", selectedKey: statusFilter, onChange: function (_, option) { return setStatusFilter((option === null || option === void 0 ? void 0 : option.key) || 'All'); }, options: statusOptions, styles: { root: { width: '200px' } } }),
            React.createElement(Dropdown, { label: "Filter by Source", selectedKey: sourceFilter, onChange: function (_, option) { return setSourceFilter((option === null || option === void 0 ? void 0 : option.key) || 'All'); }, options: sourceOptions, styles: { root: { width: '180px' } } }),
            React.createElement(PrimaryButton, { text: "\uD83D\uDD04 Refresh", onClick: loadMessages })),
        React.createElement(DetailsList, { items: filteredMessages, columns: columns, layoutMode: DetailsListLayoutMode.justified, selectionMode: SelectionMode.none, isHeaderVisible: true }),
        React.createElement(Panel, { isOpen: isPanelOpen, onDismiss: function () { return setIsPanelOpen(false); }, type: PanelType.medium, headerText: selectedMessage ? "\uD83D\uDCCA Details: ".concat(selectedMessage.Title) : '' }, selectedMessage && (React.createElement("div", { style: { padding: '10px' } },
            React.createElement(MessageBar, { messageBarType: selectedMessage.readStatus === 'Completed' ? MessageBarType.success :
                    selectedMessage.readStatus === 'In Progress' ? MessageBarType.warning :
                        MessageBarType.error },
                "Status: ",
                selectedMessage.readStatus,
                " - ",
                selectedMessage.readPercentage,
                "% read by users"),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement("h3", null, "\uD83D\uDCCB Message Details"),
                React.createElement("p", null,
                    React.createElement("strong", null, "Content:"),
                    " ",
                    selectedMessage.MessageContent || 'No content'),
                React.createElement("p", null,
                    React.createElement("strong", null, "Priority:"),
                    " ",
                    selectedMessage.Priority),
                React.createElement("p", null,
                    React.createElement("strong", null, "Target Audience:"),
                    " ",
                    selectedMessage.TargetAudience),
                React.createElement("p", null,
                    React.createElement("strong", null, "Created:"),
                    " ",
                    selectedMessage.Created.toLocaleString()),
                React.createElement("p", null,
                    React.createElement("strong", null, "Expires:"),
                    " ",
                    selectedMessage.ExpiryDate.toLocaleString())),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement("h3", null, "\uD83D\uDCCA Read Statistics"),
                React.createElement("p", null,
                    React.createElement("strong", null, "Total Reads:"),
                    " ",
                    selectedMessage.totalReads),
                React.createElement("p", null,
                    React.createElement("strong", null, "Unique Readers:"),
                    " ",
                    selectedMessage.uniqueReaders),
                React.createElement("p", null,
                    React.createElement("strong", null, "Read Percentage:"),
                    " ",
                    selectedMessage.readPercentage,
                    "%"),
                selectedMessage.lastReadDate && (React.createElement("p", null,
                    React.createElement("strong", null, "Last Read:"),
                    " ",
                    selectedMessage.lastReadDate.toLocaleString()))),
            React.createElement("div", { style: { marginTop: '20px' } },
                React.createElement("h3", null, "\uD83C\uDFAF Actions"),
                React.createElement("div", { style: { display: 'flex', gap: '10px', flexDirection: 'column' } },
                    React.createElement(PrimaryButton, { text: "\uD83D\uDCE4 Resend to Teams", onClick: function () {
                            // Implement resend functionality
                            alert('Resending message to Teams channels...');
                        } }),
                    React.createElement(DefaultButton, { text: "\uD83D\uDCE7 Send Reminder Email", onClick: function () {
                            // Implement email reminder
                            alert('Sending reminder emails to non-readers...');
                        } }),
                    React.createElement(DefaultButton, { text: "\uD83D\uDCCB Export Read Report", onClick: function () {
                            // Implement export functionality
                            alert('Exporting detailed read report...');
                        } }))))))));
};
//# sourceMappingURL=ManagerDashboard.js.map