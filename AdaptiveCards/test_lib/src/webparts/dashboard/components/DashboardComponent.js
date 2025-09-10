var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
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
import styles from './DashboardComponent.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';
import { enhancedDataService } from '../../../services/EnhancedDataService';
var DashboardComponent = /** @class */ (function (_super) {
    __extends(DashboardComponent, _super);
    function DashboardComponent(props) {
        var _this = _super.call(this, props) || this;
        _this.refreshTimer = null;
        _this.handleFilterChange = function (filterType, value) {
            _this.setState(function (prevState) {
                var _a;
                var newFilters = __assign(__assign({}, prevState.filters), (_a = {}, _a[filterType] = value, _a));
                var filteredMessages = _this.applyFilters(prevState.messages);
                return {
                    filters: newFilters,
                    filteredMessages: filteredMessages
                };
            });
        };
        _this.toggleCharts = function () {
            _this.setState(function (prevState) { return ({ showCharts: !prevState.showCharts }); });
        };
        // Navigation methods for Quick Actions
        _this.openTeamsMessageCreator = function () {
            var baseUrl = window.location.origin + window.location.pathname;
            var newUrl = "".concat(baseUrl, "?component=teams-message-creator");
            window.open(newUrl, '_blank');
        };
        _this.openManagerDashboard = function () {
            var baseUrl = window.location.origin + window.location.pathname;
            var newUrl = "".concat(baseUrl, "?component=manager-dashboard");
            window.open(newUrl, '_blank');
        };
        _this.openMessageDiagnostics = function () {
            var baseUrl = window.location.origin + window.location.pathname;
            var newUrl = "".concat(baseUrl, "?component=message-list-diagnostic");
            window.open(newUrl, '_blank');
        };
        _this.state = {
            messages: [],
            filteredMessages: [],
            loading: false,
            error: null,
            lastRefresh: null,
            customSiteUrl: '',
            filters: {
                priority: 'All',
                readStatus: 'All',
                targetAudience: 'All',
                dateRange: 'All'
            },
            showCharts: false
        };
        return _this;
    }
    DashboardComponent.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: 
                    // Initialize the enhanced data service with dataSourceUrl
                    return [4 /*yield*/, enhancedDataService.initialize(this.props.context, this.props.dataSourceUrl)];
                    case 1:
                        // Initialize the enhanced data service with dataSourceUrl
                        _a.sent();
                        // Check if we're in Teams context and handle accordingly
                        this.handleTeamsContext();
                        // Load initial data
                        this.loadMessages();
                        // Set up auto-refresh if configured
                        if (this.props.refreshInterval > 0) {
                            this.refreshTimer = setInterval(function () {
                                _this.loadMessages();
                            }, this.props.refreshInterval);
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    DashboardComponent.prototype.handleTeamsContext = function () {
        var _a, _b, _c, _d, _e;
        try {
            var url = window.location.href;
            var isTeamsUrl = url.includes('teams.microsoft.com') || url.includes('teams.office.com');
            var hasTeamsContext = ((_b = (_a = this.props.context.sdks) === null || _a === void 0 ? void 0 : _a.microsoftTeams) === null || _b === void 0 ? void 0 : _b.context) !== undefined;
            console.log('Dashboard: Teams context check:', {
                currentUrl: url,
                isTeamsUrl: isTeamsUrl,
                hasTeamsContext: hasTeamsContext,
                dataSourceUrl: this.props.dataSourceUrl
            });
            if (isTeamsUrl || hasTeamsContext) {
                console.log('Dashboard: Running in Teams context');
                // If a dataSourceUrl is configured, extract the SharePoint site URL from it
                if (this.props.dataSourceUrl && this.props.dataSourceUrl.includes('sharepoint.com')) {
                    var match = this.props.dataSourceUrl.match(/(https:\/\/[^\/]+\/[^\/]+\/[^\/]+)/);
                    if (match) {
                        var sharePointSite = match[1];
                        console.log('Dashboard: Setting SharePoint site for Teams:', sharePointSite);
                        enhancedDataService.setSharePointSiteUrl(sharePointSite);
                        this.setState({ customSiteUrl: sharePointSite });
                    }
                }
                else {
                    // Try to get SharePoint site from Teams context
                    if ((_d = (_c = this.props.context.sdks) === null || _c === void 0 ? void 0 : _c.microsoftTeams) === null || _d === void 0 ? void 0 : _d.context) {
                        var teamsContext = this.props.context.sdks.microsoftTeams.context;
                        var sharePointSite = null;
                        if ((_e = teamsContext.sharepoint) === null || _e === void 0 ? void 0 : _e.webAbsoluteUrl) {
                            sharePointSite = teamsContext.sharepoint.webAbsoluteUrl;
                        }
                        else if (teamsContext.teamSiteUrl) {
                            sharePointSite = teamsContext.teamSiteUrl;
                        }
                        if (sharePointSite) {
                            console.log('Dashboard: Using SharePoint site from Teams context:', sharePointSite);
                            enhancedDataService.setSharePointSiteUrl(sharePointSite);
                            this.setState({ customSiteUrl: sharePointSite });
                        }
                        else {
                            console.warn('Dashboard: No SharePoint site found in Teams context. Please configure dataSourceUrl in web part properties.');
                        }
                    }
                }
            }
            else {
                console.log('Dashboard: Running in SharePoint context');
            }
        }
        catch (error) {
            console.error('Dashboard: Error handling Teams context:', error);
        }
    };
    DashboardComponent.prototype.componentWillUnmount = function () {
        if (this.refreshTimer) {
            clearInterval(this.refreshTimer);
        }
    };
    DashboardComponent.prototype.componentDidUpdate = function (prevProps) {
        var _this = this;
        // Restart timer if refresh interval changed
        if (prevProps.refreshInterval !== this.props.refreshInterval) {
            if (this.refreshTimer) {
                clearInterval(this.refreshTimer);
            }
            if (this.props.refreshInterval > 0) {
                this.refreshTimer = setInterval(function () {
                    _this.loadMessages();
                }, this.props.refreshInterval);
            }
        }
    };
    DashboardComponent.prototype.isTeamsContext = function () {
        var _a, _b;
        var url = window.location.href;
        var isTeamsUrl = url.includes('teams.microsoft.com') || url.includes('teams.office.com');
        var hasTeamsContext = ((_b = (_a = this.props.context.sdks) === null || _a === void 0 ? void 0 : _a.microsoftTeams) === null || _b === void 0 ? void 0 : _b.context) !== undefined;
        return isTeamsUrl || hasTeamsContext;
    };
    DashboardComponent.prototype.loadMessages = function () {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var messages, filteredMessages, error_1;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        this.setState({ loading: true, error: null });
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 3, , 4]);
                        console.log('Dashboard: Starting to load messages...');
                        console.log('Dashboard: Current site URL:', (_c = (_b = (_a = this.props.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.web) === null || _c === void 0 ? void 0 : _c.absoluteUrl);
                        console.log('Dashboard: DataService custom site URL:', this.state.customSiteUrl);
                        return [4 /*yield*/, enhancedDataService.getMessagesForCurrentUser()];
                    case 2:
                        messages = _d.sent();
                        console.log('Dashboard: Successfully loaded messages:', messages.length);
                        filteredMessages = this.applyFilters(messages);
                        this.setState({
                            messages: messages,
                            filteredMessages: filteredMessages,
                            loading: false,
                            lastRefresh: new Date()
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _d.sent();
                        console.error('Dashboard: Error loading messages:', error_1);
                        console.error('Dashboard: Error details:', {
                            message: error_1.message,
                            stack: error_1.stack,
                            name: error_1.name
                        });
                        this.setState({
                            error: "".concat(error_1.message || 'Failed to load messages', " (Check browser console for details)"),
                            loading: false
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    DashboardComponent.prototype.handleMarkAsRead = function (messageId) {
        return __awaiter(this, void 0, void 0, function () {
            var error_2;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, enhancedDataService.markMessageAsRead(messageId)];
                    case 1:
                        _a.sent();
                        // Update the local state to reflect the read status
                        this.setState(function (prevState) { return ({
                            messages: prevState.messages.map(function (msg) {
                                return msg.Id === messageId
                                    ? __assign(__assign({}, msg), { ReadBy: (msg.ReadBy || '') + ';' + _this.props.context.pageContext.user.email }) : msg;
                            })
                        }); });
                        return [3 /*break*/, 3];
                    case 2:
                        error_2 = _a.sent();
                        console.error('Error marking message as read:', error_2);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    DashboardComponent.prototype.isMessageRead = function (message) {
        var _a;
        var currentUserEmail = this.props.context.pageContext.user.email;
        return ((_a = message.ReadBy) === null || _a === void 0 ? void 0 : _a.includes(currentUserEmail)) || false;
    };
    DashboardComponent.prototype.getPriorityColor = function (priority) {
        switch (priority) {
            case 'High': return '#d13438';
            case 'Medium': return '#ff8c00';
            case 'Low': return '#107c10';
            default: return '#605e5c';
        }
    };
    DashboardComponent.prototype.renderTitle = function () {
        var _this = this;
        if (this.props.displayMode === DisplayMode.Edit) {
            return (React.createElement("input", { type: "text", value: this.props.title, onChange: function (e) { return _this.props.updateProperty(e.target.value); }, placeholder: "Enter dashboard title", style: {
                    fontSize: '24px',
                    fontWeight: 'bold',
                    border: '1px dashed #ccc',
                    padding: '8px 12px',
                    background: 'transparent',
                    width: '100%',
                    marginBottom: '16px'
                } }));
        }
        return this.props.title ? (React.createElement("h1", { style: { marginBottom: '16px', fontSize: '24px', fontWeight: 'bold' } }, escape(this.props.title))) : null;
    };
    DashboardComponent.prototype.renderDescription = function () {
        var _this = this;
        if (!this.props.description)
            return null;
        if (this.props.displayMode === DisplayMode.Edit) {
            return (React.createElement("textarea", { value: this.props.description, onChange: function (e) { return _this.props.updateProperty(e.target.value); }, placeholder: "Enter dashboard description", style: {
                    fontSize: '14px',
                    border: '1px dashed #ccc',
                    padding: '8px 12px',
                    background: 'transparent',
                    width: '100%',
                    marginBottom: '16px',
                    minHeight: '60px',
                    resize: 'vertical'
                } }));
        }
        return (React.createElement("p", { style: { marginBottom: '16px', color: '#605e5c' } }, escape(this.props.description)));
    };
    DashboardComponent.prototype.renderRefreshInfo = function () {
        var _this = this;
        if (!this.state.lastRefresh)
            return null;
        return (React.createElement("div", { className: styles.refreshInfo },
            React.createElement("span", null,
                "Last updated: ",
                this.state.lastRefresh.toLocaleTimeString()),
            this.props.showRefreshButton && (React.createElement("button", { className: styles.refreshButton, onClick: function () { return _this.loadMessages(); }, disabled: this.state.loading },
                this.state.loading ? '⟳' : '🔄',
                " Refresh"))));
    };
    DashboardComponent.prototype.renderMessage = function (message) {
        var _this = this;
        var isRead = this.isMessageRead(message);
        var priorityColor = this.getPriorityColor(message.Priority);
        return (React.createElement("div", { key: message.Id, className: "".concat(styles.messageCard, " ").concat(isRead ? styles.readMessage : styles.unreadMessage) },
            React.createElement("div", { className: styles.messageHeader },
                React.createElement("div", { className: styles.messageTitle },
                    React.createElement("div", { className: styles.priorityIndicator, style: { backgroundColor: priorityColor } }),
                    React.createElement("h3", null, escape(message.Title))),
                React.createElement("div", { className: styles.messageMetadata },
                    React.createElement("span", { className: styles.priority, style: { color: priorityColor } }, message.Priority),
                    React.createElement("span", { className: styles.date }, new Date(message.Created).toLocaleDateString()))),
            React.createElement("div", { className: styles.messageContent },
                React.createElement("div", { dangerouslySetInnerHTML: { __html: message.MessageContent } })),
            React.createElement("div", { className: styles.messageFooter },
                React.createElement("div", { className: styles.messageInfo },
                    React.createElement("span", null,
                        "From: ",
                        escape(message.Author.Title)),
                    React.createElement("span", null,
                        "Expires: ",
                        new Date(message.ExpiryDate).toLocaleDateString())),
                !isRead && (React.createElement("button", { className: styles.markReadButton, onClick: function () { return _this.handleMarkAsRead(message.Id); } }, "Mark as Read")),
                isRead && (React.createElement("span", { className: styles.readIndicator }, "\u2713 Read")))));
    };
    DashboardComponent.prototype.renderPlaceholder = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.placeholder },
            React.createElement("div", { className: styles.placeholderIcon }, "\uD83D\uDCCA"),
            React.createElement("div", { className: styles.placeholderTitle }, "Configure your Dashboard"),
            React.createElement("div", { className: styles.placeholderDescription }, "Please configure the dashboard settings in the web part properties."),
            React.createElement("button", { className: styles.configureButton, onClick: function () { return _this.props.context.propertyPane.open(); } }, "Configure")));
    };
    // Filter and Chart Methods
    DashboardComponent.prototype.applyFilters = function (messages) {
        var _this = this;
        return messages.filter(function (message) {
            // Priority filter
            if (_this.state.filters.priority !== 'All' && message.Priority !== _this.state.filters.priority) {
                return false;
            }
            // Read status filter
            if (_this.state.filters.readStatus !== 'All') {
                var isRead = _this.isMessageRead(message);
                if (_this.state.filters.readStatus === 'Read' && !isRead)
                    return false;
                if (_this.state.filters.readStatus === 'Unread' && isRead)
                    return false;
            }
            // Target audience filter
            if (_this.state.filters.targetAudience !== 'All' && message.TargetAudience !== _this.state.filters.targetAudience) {
                return false;
            }
            // Date range filter
            if (_this.state.filters.dateRange !== 'All') {
                var now = new Date();
                var messageDate = new Date(message.Created);
                var daysDiff = Math.floor((now.getTime() - messageDate.getTime()) / (1000 * 3600 * 24));
                switch (_this.state.filters.dateRange) {
                    case 'Today':
                        if (daysDiff > 0)
                            return false;
                        break;
                    case 'This Week':
                        if (daysDiff > 7)
                            return false;
                        break;
                    case 'This Month':
                        if (daysDiff > 30)
                            return false;
                        break;
                }
            }
            return true;
        });
    };
    DashboardComponent.prototype.getChartData = function () {
        var _this = this;
        var filteredMessages = this.state.filteredMessages;
        // Priority distribution
        var priorityData = {
            High: filteredMessages.filter(function (m) { return m.Priority === 'High'; }).length,
            Medium: filteredMessages.filter(function (m) { return m.Priority === 'Medium'; }).length,
            Low: filteredMessages.filter(function (m) { return m.Priority === 'Low'; }).length
        };
        // Read status distribution
        var readData = {
            Read: filteredMessages.filter(function (m) { return _this.isMessageRead(m); }).length,
            Unread: filteredMessages.filter(function (m) { return !_this.isMessageRead(m); }).length
        };
        // Messages over time (last 7 days)
        var timeData = [];
        var _loop_1 = function (i) {
            var date = new Date();
            date.setDate(date.getDate() - i);
            var dateStr = date.toLocaleDateString();
            var count = filteredMessages.filter(function (m) {
                var msgDate = new Date(m.Created);
                return msgDate.toLocaleDateString() === dateStr;
            }).length;
            timeData.push({ date: dateStr, count: count });
        };
        for (var i = 6; i >= 0; i--) {
            _loop_1(i);
        }
        return { priorityData: priorityData, readData: readData, timeData: timeData };
    };
    DashboardComponent.prototype.renderFilters = function () {
        var _this = this;
        var filters = this.state.filters;
        var audiences = this.state.messages.map(function (m) { return m.TargetAudience; }).filter(function (value, index, self) { return self.indexOf(value) === index; });
        var uniqueAudiences = __spreadArray(['All'], audiences, true);
        return (React.createElement("div", { style: {
                display: 'flex',
                gap: '16px',
                marginBottom: '20px',
                flexWrap: 'wrap',
                alignItems: 'center'
            } },
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column' } },
                React.createElement("label", { style: { fontSize: '12px', fontWeight: '600', marginBottom: '4px' } }, "Priority"),
                React.createElement("select", { value: filters.priority, onChange: function (e) { return _this.handleFilterChange('priority', e.target.value); }, style: { padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' } },
                    React.createElement("option", { value: "All" }, "All Priorities"),
                    React.createElement("option", { value: "High" }, "High"),
                    React.createElement("option", { value: "Medium" }, "Medium"),
                    React.createElement("option", { value: "Low" }, "Low"))),
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column' } },
                React.createElement("label", { style: { fontSize: '12px', fontWeight: '600', marginBottom: '4px' } }, "Status"),
                React.createElement("select", { value: filters.readStatus, onChange: function (e) { return _this.handleFilterChange('readStatus', e.target.value); }, style: { padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' } },
                    React.createElement("option", { value: "All" }, "All Messages"),
                    React.createElement("option", { value: "Read" }, "Read"),
                    React.createElement("option", { value: "Unread" }, "Unread"))),
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column' } },
                React.createElement("label", { style: { fontSize: '12px', fontWeight: '600', marginBottom: '4px' } }, "Audience"),
                React.createElement("select", { value: filters.targetAudience, onChange: function (e) { return _this.handleFilterChange('targetAudience', e.target.value); }, style: { padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' } }, uniqueAudiences.map(function (audience) { return (React.createElement("option", { key: audience, value: audience }, audience)); }))),
            React.createElement("div", { style: { display: 'flex', flexDirection: 'column' } },
                React.createElement("label", { style: { fontSize: '12px', fontWeight: '600', marginBottom: '4px' } }, "Date Range"),
                React.createElement("select", { value: filters.dateRange, onChange: function (e) { return _this.handleFilterChange('dateRange', e.target.value); }, style: { padding: '6px 8px', border: '1px solid #ccc', borderRadius: '4px' } },
                    React.createElement("option", { value: "All" }, "All Time"),
                    React.createElement("option", { value: "Today" }, "Today"),
                    React.createElement("option", { value: "This Week" }, "This Week"),
                    React.createElement("option", { value: "This Month" }, "This Month"))),
            React.createElement("button", { onClick: this.toggleCharts, style: {
                    padding: '8px 16px',
                    background: this.state.showCharts ? '#106ebe' : '#0078d4',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: 'pointer',
                    fontSize: '12px',
                    fontWeight: '600',
                    marginTop: '18px'
                } }, this.state.showCharts ? '📊 Hide Charts' : '📈 Show Charts')));
    };
    DashboardComponent.prototype.renderCharts = function () {
        if (!this.state.showCharts)
            return null;
        var _a = this.getChartData(), priorityData = _a.priorityData, readData = _a.readData, timeData = _a.timeData;
        return (React.createElement("div", { style: { marginBottom: '20px' } },
            React.createElement("div", { style: {
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))',
                    gap: '20px',
                    marginBottom: '20px'
                } },
                React.createElement("div", { style: {
                        background: 'white',
                        padding: '20px',
                        borderRadius: '8px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                    } },
                    React.createElement("h3", { style: { marginTop: 0, marginBottom: '16px', fontSize: '16px' } }, "Priority Distribution"),
                    React.createElement("div", { style: { height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'space-around' } },
                        React.createElement("div", { style: { textAlign: 'center' } },
                            React.createElement("div", { style: {
                                    width: '60px',
                                    height: '60px',
                                    background: '#d13438',
                                    borderRadius: '50%',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    color: 'white',
                                    fontSize: '18px',
                                    fontWeight: 'bold',
                                    margin: '0 auto 8px auto'
                                } }, priorityData.High),
                            React.createElement("div", { style: { fontSize: '12px' } }, "High")),
                        React.createElement("div", { style: { textAlign: 'center' } },
                            React.createElement("div", { style: {
                                    width: '60px',
                                    height: '60px',
                                    background: '#ff8c00',
                                    borderRadius: '50%',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    color: 'white',
                                    fontSize: '18px',
                                    fontWeight: 'bold',
                                    margin: '0 auto 8px auto'
                                } }, priorityData.Medium),
                            React.createElement("div", { style: { fontSize: '12px' } }, "Medium")),
                        React.createElement("div", { style: { textAlign: 'center' } },
                            React.createElement("div", { style: {
                                    width: '60px',
                                    height: '60px',
                                    background: '#107c10',
                                    borderRadius: '50%',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    color: 'white',
                                    fontSize: '18px',
                                    fontWeight: 'bold',
                                    margin: '0 auto 8px auto'
                                } }, priorityData.Low),
                            React.createElement("div", { style: { fontSize: '12px' } }, "Low")))),
                React.createElement("div", { style: {
                        background: 'white',
                        padding: '20px',
                        borderRadius: '8px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                    } },
                    React.createElement("h3", { style: { marginTop: 0, marginBottom: '16px', fontSize: '16px' } }, "Read Status"),
                    React.createElement("div", { style: { height: '200px', display: 'flex', alignItems: 'center', justifyContent: 'space-around' } },
                        React.createElement("div", { style: { textAlign: 'center' } },
                            React.createElement("div", { style: {
                                    width: '80px',
                                    height: "".concat(Math.max(20, (readData.Read / Math.max(readData.Read + readData.Unread, 1)) * 150), "px"),
                                    background: '#107c10',
                                    margin: '0 auto 8px auto',
                                    display: 'flex',
                                    alignItems: 'flex-end',
                                    justifyContent: 'center',
                                    color: 'white',
                                    fontWeight: 'bold',
                                    paddingBottom: '8px'
                                } }, readData.Read),
                            React.createElement("div", { style: { fontSize: '12px' } }, "Read")),
                        React.createElement("div", { style: { textAlign: 'center' } },
                            React.createElement("div", { style: {
                                    width: '80px',
                                    height: "".concat(Math.max(20, (readData.Unread / Math.max(readData.Read + readData.Unread, 1)) * 150), "px"),
                                    background: '#d13438',
                                    margin: '0 auto 8px auto',
                                    display: 'flex',
                                    alignItems: 'flex-end',
                                    justifyContent: 'center',
                                    color: 'white',
                                    fontWeight: 'bold',
                                    paddingBottom: '8px'
                                } }, readData.Unread),
                            React.createElement("div", { style: { fontSize: '12px' } }, "Unread"))))),
            React.createElement("div", { style: {
                    background: 'white',
                    padding: '20px',
                    borderRadius: '8px',
                    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                } },
                React.createElement("h3", { style: { marginTop: 0, marginBottom: '16px', fontSize: '16px' } }, "Messages Over Time (Last 7 Days)"),
                React.createElement("div", { style: {
                        height: '150px',
                        display: 'flex',
                        alignItems: 'flex-end',
                        justifyContent: 'space-between',
                        borderBottom: '1px solid #ccc',
                        paddingBottom: '10px'
                    } }, timeData.map(function (day, index) { return (React.createElement("div", { key: index, style: { textAlign: 'center', flex: 1 } },
                    React.createElement("div", { style: {
                            height: "".concat(Math.max(10, (day.count / Math.max.apply(Math, __spreadArray(__spreadArray([], timeData.map(function (d) { return d.count; }), false), [1], false))) * 100), "px"),
                            background: '#0078d4',
                            margin: '0 auto 8px auto',
                            width: '30px',
                            display: 'flex',
                            alignItems: 'flex-end',
                            justifyContent: 'center',
                            color: 'white',
                            fontSize: '12px',
                            fontWeight: 'bold',
                            paddingBottom: '4px'
                        } }, day.count),
                    React.createElement("div", { style: { fontSize: '10px', transform: 'rotate(-45deg)', transformOrigin: 'center' } }, day.date.split('/').slice(0, 2).join('/')))); })))));
    };
    DashboardComponent.prototype.render = function () {
        var _this = this;
        var _a = this.state, loading = _a.loading, error = _a.error, messages = _a.messages, filteredMessages = _a.filteredMessages, showCharts = _a.showCharts;
        return (React.createElement("div", { className: styles.dashboardComponent },
            React.createElement("div", { style: { marginBottom: '20px' } },
                React.createElement("h2", null, "\uD83D\uDCCA Personal Dashboard"),
                React.createElement("p", null, "Monitor your personalized messages and activity")),
            React.createElement("div", { style: { marginBottom: '20px', padding: '16px', background: '#e8f4fd', borderRadius: '8px', border: '1px solid #0078d4' } },
                React.createElement("h3", { style: { margin: '0 0 12px 0', fontSize: '16px', color: '#323130' } }, "\uD83D\uDE80 Quick Actions"),
                React.createElement("div", { style: { display: 'flex', gap: '12px', flexWrap: 'wrap' } },
                    React.createElement("button", { onClick: function () { return _this.openTeamsMessageCreator(); }, style: {
                            padding: '8px 16px',
                            backgroundColor: '#0078d4',
                            color: 'white',
                            border: 'none',
                            borderRadius: '4px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600'
                        } }, "\uD83D\uDCDD Create New Message"),
                    React.createElement("button", { onClick: function () { return _this.openManagerDashboard(); }, style: {
                            padding: '8px 16px',
                            backgroundColor: '#107c10',
                            color: 'white',
                            border: 'none',
                            borderRadius: '4px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600'
                        } }, "\uD83D\uDC65 Manager Dashboard"),
                    React.createElement("button", { onClick: function () { return _this.openMessageDiagnostics(); }, style: {
                            padding: '8px 16px',
                            backgroundColor: '#ca5010',
                            color: 'white',
                            border: 'none',
                            borderRadius: '4px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600'
                        } }, "\uD83D\uDD0D Message Diagnostics"))),
            React.createElement("div", { style: { marginBottom: '20px', padding: '16px', background: '#f8f9fa', borderRadius: '8px', border: '1px solid #e1e5e9' } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', marginBottom: '12px' } },
                    React.createElement("h3", { style: { margin: 0, fontSize: '16px', color: '#323130' } }, "\u2699\uFE0F Data Source Configuration")),
                React.createElement("div", { style: { marginBottom: '12px' } },
                    React.createElement("label", { style: { display: 'block', fontSize: '14px', fontWeight: '600', marginBottom: '4px' } }, "SharePoint Site URL:"),
                    React.createElement("input", { type: "text", value: this.state.customSiteUrl || '', onChange: function (e) { return _this.setState({ customSiteUrl: e.target.value }); }, placeholder: "https://yourtenant.sharepoint.com/sites/yoursite (leave empty to use current site)", style: {
                            width: '100%',
                            padding: '8px 12px',
                            border: '1px solid #d1d1d1',
                            borderRadius: '4px',
                            fontSize: '14px'
                        } })),
                React.createElement("div", { style: { fontSize: '12px', color: '#605e5c', lineHeight: '1.4' } },
                    React.createElement("strong", null, "Usage:"),
                    " Configure which SharePoint site to load data from. Leave empty to use the current site where this web part is deployed. This allows you to centralize your Adaptive Cards data in one location while deploying the Dashboard web part to multiple sites.")),
            React.createElement("div", { style: { marginBottom: '20px', padding: '12px', background: '#f3f2f1', borderRadius: '6px', border: '1px solid #edebe9' } },
                React.createElement("div", { style: { display: 'flex', alignItems: 'center', fontSize: '14px' } },
                    React.createElement("span", { style: { marginRight: '8px' } }, this.isTeamsContext() ? '👥' : '📋'),
                    React.createElement("span", { style: { fontWeight: '600' } }, this.isTeamsContext() ? 'Teams Context' : 'SharePoint Context'),
                    this.state.customSiteUrl && (React.createElement("span", { style: { marginLeft: '12px', color: '#605e5c' } },
                        "\u2192 ",
                        this.state.customSiteUrl)))),
            !loading && this.renderFilters(),
            !loading && this.renderCharts(),
            loading && (React.createElement("div", { className: styles.loading },
                React.createElement("div", { className: styles.spinner }),
                React.createElement("span", null, "Loading messages..."))),
            error && (React.createElement("div", { className: styles.error },
                React.createElement("div", { className: styles.errorIcon }, "\u26A0\uFE0F"),
                React.createElement("div", null,
                    React.createElement("strong", null, "Error loading dashboard:"),
                    React.createElement("br", null),
                    error,
                    React.createElement("div", { style: { fontSize: '14px', marginTop: '8px' } },
                        React.createElement("strong", null, "Note:"),
                        " Dashboard is showing sample data below. To connect to live SharePoint data:",
                        React.createElement("ol", { style: { marginTop: '8px', paddingLeft: '20px' } },
                            React.createElement("li", null, "Ensure the SharePoint lists exist (run setup-sharepoint-lists.ps1)"),
                            React.createElement("li", null, "Verify permissions to access the configured site"),
                            React.createElement("li", null, "Check that the site URL is correct")))))),
            React.createElement("div", { style: {
                    display: 'grid',
                    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
                    gap: '16px',
                    marginBottom: '24px'
                } },
                React.createElement("div", { style: {
                        background: 'white',
                        padding: '20px',
                        borderRadius: '8px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                        textAlign: 'center'
                    } },
                    React.createElement("div", { style: { fontSize: '32px', marginBottom: '8px' } }, "\uD83D\uDCEC"),
                    React.createElement("div", { style: { fontSize: '24px', fontWeight: 'bold', color: '#0078d4' } }, filteredMessages.length),
                    React.createElement("div", { style: { fontSize: '12px', color: '#666' } }, "Total Messages")),
                React.createElement("div", { style: {
                        background: 'white',
                        padding: '20px',
                        borderRadius: '8px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                        textAlign: 'center'
                    } },
                    React.createElement("div", { style: { fontSize: '32px', marginBottom: '8px' } }, "\u2705"),
                    React.createElement("div", { style: { fontSize: '24px', fontWeight: 'bold', color: '#107c10' } }, filteredMessages.filter(function (m) { return _this.isMessageRead(m); }).length),
                    React.createElement("div", { style: { fontSize: '12px', color: '#666' } }, "Read Messages")),
                React.createElement("div", { style: {
                        background: 'white',
                        padding: '20px',
                        borderRadius: '8px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                        textAlign: 'center'
                    } },
                    React.createElement("div", { style: { fontSize: '32px', marginBottom: '8px' } }, "\uD83D\uDD14"),
                    React.createElement("div", { style: { fontSize: '24px', fontWeight: 'bold', color: '#d13438' } }, filteredMessages.filter(function (m) { return !_this.isMessageRead(m); }).length),
                    React.createElement("div", { style: { fontSize: '12px', color: '#666' } }, "Unread Messages")),
                React.createElement("div", { style: {
                        background: 'white',
                        padding: '20px',
                        borderRadius: '8px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
                        textAlign: 'center'
                    } },
                    React.createElement("div", { style: { fontSize: '32px', marginBottom: '8px' } }, "\u26A1"),
                    React.createElement("div", { style: { fontSize: '24px', fontWeight: 'bold', color: '#ff8c00' } }, filteredMessages.filter(function (m) { return m.Priority === 'High'; }).length),
                    React.createElement("div", { style: { fontSize: '12px', color: '#666' } }, "High Priority"))),
            React.createElement("div", { className: styles.messagesContainer },
                React.createElement("div", { className: styles.messagesHeader },
                    React.createElement("h2", null, "\uD83D\uDCCB Your Messages"),
                    React.createElement("div", { className: styles.messageStats },
                        React.createElement("span", null,
                            "Showing ",
                            filteredMessages.length,
                            " messages"))),
                filteredMessages.length === 0 ? (React.createElement("div", { className: styles.noMessages },
                    React.createElement("div", { className: styles.noMessagesIcon }, "\uD83D\uDCED"),
                    React.createElement("h3", null, "No messages found"),
                    React.createElement("p", null, error ? 'Try adjusting your filters or check your data connection.' : 'You\'re all caught up!'))) : (React.createElement("div", { className: styles.messagesList }, filteredMessages.map(function (message) { return _this.renderMessage(message); })))),
            React.createElement("div", { style: {
                    marginTop: '30px',
                    padding: '20px',
                    background: 'white',
                    borderRadius: '8px',
                    boxShadow: '0 2px 4px rgba(0,0,0,0.1)'
                } },
                React.createElement("h3", { style: {
                        marginBottom: '16px',
                        color: '#323130',
                        fontSize: '16px',
                        fontWeight: '600'
                    } }, "\uD83D\uDE80 Quick Actions"),
                React.createElement("div", { style: {
                        display: 'flex',
                        gap: '12px',
                        flexWrap: 'wrap'
                    } },
                    React.createElement("button", { onClick: this.openTeamsMessageCreator, style: {
                            padding: '12px 20px',
                            background: '#0078d4',
                            color: 'white',
                            border: 'none',
                            borderRadius: '6px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            transition: 'all 0.2s ease'
                        }, onMouseOver: function (e) {
                            e.currentTarget.style.background = '#106ebe';
                            e.currentTarget.style.transform = 'translateY(-1px)';
                        }, onMouseOut: function (e) {
                            e.currentTarget.style.background = '#0078d4';
                            e.currentTarget.style.transform = 'translateY(0)';
                        } }, "\uD83D\uDCDD Create Teams Message"),
                    React.createElement("button", { onClick: this.openManagerDashboard, style: {
                            padding: '12px 20px',
                            background: '#107c10',
                            color: 'white',
                            border: 'none',
                            borderRadius: '6px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            transition: 'all 0.2s ease'
                        }, onMouseOver: function (e) {
                            e.currentTarget.style.background = '#0e6e0e';
                            e.currentTarget.style.transform = 'translateY(-1px)';
                        }, onMouseOut: function (e) {
                            e.currentTarget.style.background = '#107c10';
                            e.currentTarget.style.transform = 'translateY(0)';
                        } }, "\uD83C\uDF9B\uFE0F Manager Dashboard"),
                    React.createElement("button", { onClick: this.openMessageDiagnostics, style: {
                            padding: '12px 20px',
                            background: '#d13438',
                            color: 'white',
                            border: 'none',
                            borderRadius: '6px',
                            cursor: 'pointer',
                            fontSize: '14px',
                            fontWeight: '600',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            transition: 'all 0.2s ease'
                        }, onMouseOver: function (e) {
                            e.currentTarget.style.background = '#b92b2b';
                            e.currentTarget.style.transform = 'translateY(-1px)';
                        }, onMouseOut: function (e) {
                            e.currentTarget.style.background = '#d13438';
                            e.currentTarget.style.transform = 'translateY(0)';
                        } }, "\uD83D\uDD0D Message Diagnostics")))));
    };
    return DashboardComponent;
}(React.Component));
export { DashboardComponent };
//# sourceMappingURL=DashboardComponent.js.map