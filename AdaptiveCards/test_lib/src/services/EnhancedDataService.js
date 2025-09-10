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
import { SPHttpClient } from '@microsoft/sp-http';
import { graphService } from './GraphService';
import { ManagersListService } from './ManagersListService';
var EnhancedDataService = /** @class */ (function () {
    function EnhancedDataService() {
        this.MESSAGES_LIST = 'Important Messages';
        this.READ_ACTIONS_LIST = 'MessageReadConfirmations';
        this.customSiteUrl = '';
        this.currentUser = null;
        this.mapToMessage = function (item) {
            return {
                Id: item.Id,
                Title: item.Title,
                MessageContent: item.MessageContent,
                Priority: item.Priority,
                ExpiryDate: item.ExpiryDate ? new Date(item.ExpiryDate) : undefined,
                TargetAudience: item.TargetAudience,
                ReadBy: item.ReadBy,
                ReadById: item.ReadBy ? item.ReadBy.split(';').map(function (email) { return email.trim(); }).filter(function (email) { return email; }) : [],
                Created: new Date(item.Created),
                Modified: new Date(item.Modified),
                Author: {
                    Title: 'System User',
                    Email: 'system@company.com'
                }
            };
        };
    }
    /**
     * Initialize the service with SharePoint context and Graph service
     */
    EnhancedDataService.prototype.initialize = function (context, dataSourceUrl) {
        var _a, _b, _c, _d, _e;
        return __awaiter(this, void 0, void 0, function () {
            var error_1, _f, error_2, match;
            return __generator(this, function (_g) {
                switch (_g.label) {
                    case 0:
                        this.context = context;
                        this.managersService = new ManagersListService(context);
                        _g.label = 1;
                    case 1:
                        _g.trys.push([1, 3, , 4]);
                        // Initialize Graph service with proper error handling
                        console.log('EnhancedDataService: Initializing Graph service...');
                        return [4 /*yield*/, graphService.initialize(context)];
                    case 2:
                        _g.sent();
                        console.log('EnhancedDataService: Graph service initialized successfully');
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _g.sent();
                        console.warn('EnhancedDataService: Graph service initialization failed, continuing without Graph:', error_1);
                        return [3 /*break*/, 4];
                    case 4:
                        _g.trys.push([4, 6, , 7]);
                        // Get enhanced user information with fallback
                        console.log('EnhancedDataService: Getting enhanced user information...');
                        _f = this;
                        return [4 /*yield*/, this.getEnhancedCurrentUser()];
                    case 5:
                        _f.currentUser = _g.sent();
                        console.log('EnhancedDataService: Enhanced user information retrieved');
                        return [3 /*break*/, 7];
                    case 6:
                        error_2 = _g.sent();
                        console.warn('EnhancedDataService: Failed to get enhanced user info, using basic fallback:', error_2);
                        this.currentUser = {
                            displayName: ((_b = (_a = context === null || context === void 0 ? void 0 : context.pageContext) === null || _a === void 0 ? void 0 : _a.user) === null || _b === void 0 ? void 0 : _b.displayName) || 'Unknown',
                            email: ((_d = (_c = context === null || context === void 0 ? void 0 : context.pageContext) === null || _c === void 0 ? void 0 : _c.user) === null || _d === void 0 ? void 0 : _d.email) || '',
                            graph: null,
                            spfx: ((_e = context === null || context === void 0 ? void 0 : context.pageContext) === null || _e === void 0 ? void 0 : _e.user) ? {
                                displayName: context.pageContext.user.displayName,
                                email: context.pageContext.user.email,
                                loginName: context.pageContext.user.loginName
                            } : null,
                            groups: [],
                            hasPhoto: false,
                            isManager: false,
                            isAdmin: false
                        };
                        return [3 /*break*/, 7];
                    case 7:
                        // Set custom site URL if provided
                        if (dataSourceUrl && dataSourceUrl.includes('sharepoint.com')) {
                            match = dataSourceUrl.match(/(https:\/\/[^\/]+\/[^\/]+\/[^\/]+)/);
                            if (match) {
                                this.customSiteUrl = match[1];
                                console.log("EnhancedDataService: Using custom site URL: ".concat(this.customSiteUrl));
                            }
                        }
                        console.log('EnhancedDataService: Initialized successfully');
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get enhanced user information combining Graph API and SPFx context
     */
    EnhancedDataService.prototype.getEnhancedCurrentUser = function () {
        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
        return __awaiter(this, void 0, void 0, function () {
            var enhanced, userEmail, isManager, error_3, error_4, email, isManager, error_5;
            return __generator(this, function (_o) {
                switch (_o.label) {
                    case 0:
                        _o.trys.push([0, 7, , 13]);
                        return [4 /*yield*/, graphService.getEnhancedUserInfo()];
                    case 1:
                        enhanced = _o.sent();
                        userEmail = ((_a = enhanced.graph) === null || _a === void 0 ? void 0 : _a.mail) || ((_b = enhanced.context) === null || _b === void 0 ? void 0 : _b.email) || '';
                        isManager = false;
                        _o.label = 2;
                    case 2:
                        _o.trys.push([2, 5, , 6]);
                        if (!(this.managersService && userEmail)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.managersService.isUserManager(userEmail)];
                    case 3:
                        isManager = _o.sent();
                        _o.label = 4;
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        error_3 = _o.sent();
                        console.warn('EnhancedDataService: Error checking manager status from SharePoint list:', error_3);
                        // Fallback to Graph service result
                        isManager = enhanced.isManager;
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/, {
                            displayName: ((_c = enhanced.graph) === null || _c === void 0 ? void 0 : _c.displayName) || ((_d = enhanced.context) === null || _d === void 0 ? void 0 : _d.displayName) || 'Unknown',
                            email: userEmail,
                            graph: enhanced.graph,
                            spfx: enhanced.context,
                            groups: enhanced.groups,
                            hasPhoto: !!enhanced.photo,
                            isManager: isManager,
                            isAdmin: enhanced.isAdmin
                        }];
                    case 7:
                        error_4 = _o.sent();
                        console.warn('EnhancedDataService: Error getting enhanced user info, using SPFx fallback:', error_4);
                        email = ((_g = (_f = (_e = this.context) === null || _e === void 0 ? void 0 : _e.pageContext) === null || _f === void 0 ? void 0 : _f.user) === null || _g === void 0 ? void 0 : _g.email) || '';
                        isManager = false;
                        _o.label = 8;
                    case 8:
                        _o.trys.push([8, 11, , 12]);
                        if (!(this.managersService && email)) return [3 /*break*/, 10];
                        return [4 /*yield*/, this.managersService.isUserManager(email)];
                    case 9:
                        isManager = _o.sent();
                        _o.label = 10;
                    case 10: return [3 /*break*/, 12];
                    case 11:
                        error_5 = _o.sent();
                        console.warn('EnhancedDataService: Error checking manager status in fallback:', error_5);
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/, {
                            displayName: ((_k = (_j = (_h = this.context) === null || _h === void 0 ? void 0 : _h.pageContext) === null || _j === void 0 ? void 0 : _j.user) === null || _k === void 0 ? void 0 : _k.displayName) || 'Unknown',
                            email: email,
                            graph: null,
                            spfx: ((_m = (_l = this.context) === null || _l === void 0 ? void 0 : _l.pageContext) === null || _m === void 0 ? void 0 : _m.user) ? {
                                displayName: this.context.pageContext.user.displayName,
                                email: this.context.pageContext.user.email,
                                loginName: this.context.pageContext.user.loginName
                            } : null,
                            groups: [],
                            hasPhoto: false,
                            isManager: isManager,
                            isAdmin: false
                        }];
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if the current user is a manager according to SharePoint Managers list
     */
    EnhancedDataService.prototype.isCurrentUserManager = function () {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var userEmail, error_6;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _d.trys.push([0, 2, , 3]);
                        if (!this.managersService) {
                            console.warn('EnhancedDataService: ManagersListService not initialized');
                            return [2 /*return*/, false];
                        }
                        userEmail = (_c = (_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.user) === null || _c === void 0 ? void 0 : _c.email;
                        if (!userEmail) {
                            console.warn('EnhancedDataService: No user email available');
                            return [2 /*return*/, false];
                        }
                        return [4 /*yield*/, this.managersService.isUserManager(userEmail)];
                    case 1: return [2 /*return*/, _d.sent()];
                    case 2:
                        error_6 = _d.sent();
                        console.error('EnhancedDataService: Error checking manager status:', error_6);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if a specific user is a manager
     */
    EnhancedDataService.prototype.isUserManager = function (userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        if (!this.managersService) {
                            console.warn('EnhancedDataService: ManagersListService not initialized');
                            return [2 /*return*/, false];
                        }
                        return [4 /*yield*/, this.managersService.isUserManager(userEmail)];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        error_7 = _a.sent();
                        console.error('EnhancedDataService: Error checking manager status for user:', userEmail, error_7);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get manager details for the current user
     */
    EnhancedDataService.prototype.getCurrentUserManagerDetails = function () {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var userEmail, error_8;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _d.trys.push([0, 2, , 3]);
                        if (!this.managersService) {
                            console.warn('EnhancedDataService: ManagersListService not initialized');
                            return [2 /*return*/, null];
                        }
                        userEmail = (_c = (_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.user) === null || _c === void 0 ? void 0 : _c.email;
                        if (!userEmail) {
                            console.warn('EnhancedDataService: No user email available');
                            return [2 /*return*/, null];
                        }
                        return [4 /*yield*/, this.managersService.getManagerDetails(userEmail)];
                    case 1: return [2 /*return*/, _d.sent()];
                    case 2:
                        error_8 = _d.sent();
                        console.error('EnhancedDataService: Error getting manager details:', error_8);
                        return [2 /*return*/, null];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get all active managers from SharePoint list
     */
    EnhancedDataService.prototype.getAllManagers = function () {
        return __awaiter(this, void 0, void 0, function () {
            var error_9;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        if (!this.managersService) {
                            console.warn('EnhancedDataService: ManagersListService not initialized');
                            return [2 /*return*/, []];
                        }
                        return [4 /*yield*/, this.managersService.getActiveManagers()];
                    case 1: return [2 /*return*/, _a.sent()];
                    case 2:
                        error_9 = _a.sent();
                        console.error('EnhancedDataService: Error getting all managers:', error_9);
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Set a custom SharePoint site URL
     */
    EnhancedDataService.prototype.setSharePointSiteUrl = function (siteUrl) {
        var normalizedUrl = siteUrl.replace(/\/$/, '');
        console.log("EnhancedDataService: Setting custom SharePoint site URL: ".concat(normalizedUrl));
        this.customSiteUrl = normalizedUrl;
    };
    /**
     * Get the current site URL for API calls
     */
    EnhancedDataService.prototype.getCurrentSiteUrl = function () {
        var _a, _b, _c, _d, _e, _f;
        var siteUrl;
        if (this.isTeamsContext()) {
            if (this.customSiteUrl) {
                siteUrl = this.customSiteUrl;
            }
            else {
                var teamsSiteUrl = this.getSharePointSiteFromTeamsContext();
                if (teamsSiteUrl) {
                    siteUrl = teamsSiteUrl;
                }
                else {
                    console.warn('EnhancedDataService: Teams context detected but no SharePoint site configured.');
                    siteUrl = ((_c = (_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.web) === null || _c === void 0 ? void 0 : _c.absoluteUrl) || 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken';
                }
            }
        }
        else {
            siteUrl = this.customSiteUrl || ((_f = (_e = (_d = this.context) === null || _d === void 0 ? void 0 : _d.pageContext) === null || _e === void 0 ? void 0 : _e.web) === null || _f === void 0 ? void 0 : _f.absoluteUrl) || 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken';
        }
        return siteUrl.replace(/\/$/, '');
    };
    /**
     * Check if we're running in Teams context
     */
    EnhancedDataService.prototype.isTeamsContext = function () {
        var _a, _b, _c;
        var url = window.location.href;
        var isTeamsUrl = url.includes('teams.microsoft.com') || url.includes('teams.office.com');
        var hasTeamsContext = false;
        try {
            hasTeamsContext = ((_c = (_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.sdks) === null || _b === void 0 ? void 0 : _b.microsoftTeams) === null || _c === void 0 ? void 0 : _c.context) !== undefined;
        }
        catch (error) {
            hasTeamsContext = false;
        }
        return isTeamsUrl || hasTeamsContext;
    };
    /**
     * Try to get SharePoint site URL from Teams context
     */
    EnhancedDataService.prototype.getSharePointSiteFromTeamsContext = function () {
        var _a, _b, _c, _d, _e, _f, _g, _h;
        try {
            if ((_c = (_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.sdks) === null || _b === void 0 ? void 0 : _b.microsoftTeams) === null || _c === void 0 ? void 0 : _c.context) {
                var teamsContext = this.context.sdks.microsoftTeams.context;
                if ((_d = teamsContext.sharepoint) === null || _d === void 0 ? void 0 : _d.serverRelativeUrl) {
                    var currentUrl = (_g = (_f = (_e = this.context) === null || _e === void 0 ? void 0 : _e.pageContext) === null || _f === void 0 ? void 0 : _f.web) === null || _g === void 0 ? void 0 : _g.absoluteUrl;
                    if (currentUrl) {
                        var tenant = currentUrl.split('/')[2];
                        return "https://".concat(tenant).concat(teamsContext.sharepoint.serverRelativeUrl);
                    }
                }
                if (teamsContext.teamSiteUrl) {
                    return teamsContext.teamSiteUrl;
                }
                if ((_h = teamsContext.sharepoint) === null || _h === void 0 ? void 0 : _h.webAbsoluteUrl) {
                    return teamsContext.sharepoint.webAbsoluteUrl;
                }
            }
            return null;
        }
        catch (error) {
            console.warn('EnhancedDataService: Error getting SharePoint site from Teams context:', error);
            return null;
        }
    };
    /**
     * Get all active messages with enhanced user filtering
     */
    EnhancedDataService.prototype.getActiveMessages = function () {
        return __awaiter(this, void 0, void 0, function () {
            var restUrl, response, data, messages, error_10;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        restUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/items") +
                            "?$select=Id,Title,MessageContent,Priority,TargetAudience,ReadBy,Created,Modified" +
                            "&$orderby=Priority desc,Created desc";
                        return [4 /*yield*/, this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP ".concat(response.status, ": ").concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        messages = data.value.map(this.mapToMessage);
                        // Filter messages based on user's groups and role
                        return [2 /*return*/, this.filterMessagesForCurrentUser(messages)];
                    case 3:
                        error_10 = _a.sent();
                        console.error('EnhancedDataService: Error fetching active messages:', error_10);
                        return [2 /*return*/, this.getMockMessages()];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Enhanced message filtering based on user's Graph groups and properties
     */
    EnhancedDataService.prototype.filterMessagesForCurrentUser = function (messages) {
        var _a, _b, _c, _d;
        if (!this.currentUser) {
            return messages; // Return all if no user context
        }
        var userGroups = this.currentUser.groups || [];
        var userEmail = ((_a = this.currentUser.spfx) === null || _a === void 0 ? void 0 : _a.email) || ((_b = this.currentUser.graph) === null || _b === void 0 ? void 0 : _b.mail) || '';
        var userDepartment = ((_c = this.currentUser.graph) === null || _c === void 0 ? void 0 : _c.department) || '';
        var userJobTitle = ((_d = this.currentUser.graph) === null || _d === void 0 ? void 0 : _d.jobTitle) || '';
        return messages.filter(function (message) {
            var targetAudience = message.TargetAudience || '';
            // Always show messages for all users
            if (targetAudience === 'All Users' || targetAudience === 'Alla Medarbetare') {
                return true;
            }
            // Check if user's groups match target audience
            var matchesGroup = userGroups.some(function (group) {
                return targetAudience.toLowerCase().includes(group.toLowerCase());
            });
            // Check if user's department matches
            var matchesDepartment = userDepartment &&
                targetAudience.toLowerCase().includes(userDepartment.toLowerCase());
            // Check if user's job title matches
            var matchesJobTitle = userJobTitle &&
                targetAudience.toLowerCase().includes(userJobTitle.toLowerCase());
            // Check if user's email is specifically mentioned
            var matchesEmail = targetAudience.toLowerCase().includes(userEmail.toLowerCase());
            console.log("EnhancedDataService: Message \"".concat(message.Title, "\" - Target: ").concat(targetAudience, ", User Groups: [").concat(userGroups.join(', '), "], Match: ").concat(matchesGroup || matchesDepartment || matchesJobTitle || matchesEmail));
            return matchesGroup || matchesDepartment || matchesJobTitle || matchesEmail;
        });
    };
    /**
     * Get messages for current user with enhanced targeting
     */
    EnhancedDataService.prototype.getMessagesForCurrentUser = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _a, siteUrl, allMessages, filteredMessages, error_11;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 4, , 5]);
                        if (!!this.currentUser) return [3 /*break*/, 2];
                        _a = this;
                        return [4 /*yield*/, this.getEnhancedCurrentUser()];
                    case 1:
                        _a.currentUser = _b.sent();
                        _b.label = 2;
                    case 2:
                        // Check if we're in development mode
                        if (!this.context || !this.context.pageContext || !this.context.pageContext.web) {
                            console.log('EnhancedDataService: Development mode detected - returning mock data');
                            return [2 /*return*/, this.getMockMessages()];
                        }
                        siteUrl = this.getCurrentSiteUrl();
                        if (!siteUrl) {
                            console.warn('EnhancedDataService: No site URL available - returning mock data');
                            return [2 /*return*/, this.getMockMessages()];
                        }
                        return [4 /*yield*/, this.getMessagesWithProgressiveQuerying(siteUrl)];
                    case 3:
                        allMessages = _b.sent();
                        if (!allMessages || allMessages.length === 0) {
                            console.warn('EnhancedDataService: No messages returned - returning mock data');
                            return [2 /*return*/, this.getMockMessages()];
                        }
                        filteredMessages = this.filterMessagesForCurrentUser(allMessages);
                        console.log("EnhancedDataService: Found ".concat(allMessages.length, " total messages, ").concat(filteredMessages.length, " for current user"));
                        return [2 /*return*/, filteredMessages];
                    case 4:
                        error_11 = _b.sent();
                        console.error('EnhancedDataService: Error fetching messages for current user:', error_11);
                        return [2 /*return*/, this.getMockMessages()];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Mark message as read with enhanced user tracking
     */
    EnhancedDataService.prototype.markMessageAsRead = function (messageId) {
        var _a, _b, _c, _d, _e;
        return __awaiter(this, void 0, void 0, function () {
            var currentUser, alreadyRead, readAction, restUrl, spOpts, _f, response, error_12;
            var _g, _h;
            return __generator(this, function (_j) {
                switch (_j.label) {
                    case 0:
                        _j.trys.push([0, 5, , 6]);
                        currentUser = ((_a = this.currentUser) === null || _a === void 0 ? void 0 : _a.spfx) || this.context.pageContext.user;
                        if (!currentUser) {
                            throw new Error('No user context available');
                        }
                        return [4 /*yield*/, this.hasUserReadMessage(messageId)];
                    case 1:
                        alreadyRead = _j.sent();
                        if (alreadyRead) {
                            console.log("Message ".concat(messageId, " already marked as read by user ").concat(currentUser.email));
                            return [2 /*return*/];
                        }
                        readAction = {
                            Title: "Read action for message ".concat(messageId),
                            MessageId: messageId,
                            UserId: parseInt(currentUser.loginName.split('|')[2] || '0'),
                            UserEmail: currentUser.email,
                            UserDisplayName: currentUser.displayName,
                            ReadTimestamp: new Date().toISOString(),
                            DeviceInfo: this.getDeviceInfo(),
                            UserDepartment: ((_c = (_b = this.currentUser) === null || _b === void 0 ? void 0 : _b.graph) === null || _c === void 0 ? void 0 : _c.department) || '',
                            UserJobTitle: ((_e = (_d = this.currentUser) === null || _d === void 0 ? void 0 : _d.graph) === null || _e === void 0 ? void 0 : _e.jobTitle) || ''
                        };
                        restUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('").concat(this.READ_ACTIONS_LIST, "')/items");
                        _g = {};
                        _h = {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-type': 'application/json;odata=nometadata',
                            'odata-version': ''
                        };
                        _f = 'X-RequestDigest';
                        return [4 /*yield*/, this.getRequestDigest()];
                    case 2:
                        spOpts = (_g.headers = (_h[_f] = _j.sent(),
                            _h),
                            _g.body = JSON.stringify(readAction),
                            _g);
                        return [4 /*yield*/, this.context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOpts)];
                    case 3:
                        response = _j.sent();
                        if (!response.ok) {
                            throw new Error("Failed to create read action: ".concat(response.status, " ").concat(response.statusText));
                        }
                        // Update the ReadBy field in the main message
                        return [4 /*yield*/, this.updateMessageReadBy(messageId, currentUser.email)];
                    case 4:
                        // Update the ReadBy field in the main message
                        _j.sent();
                        console.log("Message ".concat(messageId, " marked as read by ").concat(currentUser.email));
                        return [3 /*break*/, 6];
                    case 5:
                        error_12 = _j.sent();
                        console.error("EnhancedDataService: Error marking message ".concat(messageId, " as read:"), error_12);
                        throw new Error('Failed to mark message as read');
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if current user has read a specific message
     */
    EnhancedDataService.prototype.hasUserReadMessage = function (messageId) {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var currentUser, restUrl, response, data, error_13;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 3, , 4]);
                        currentUser = ((_a = this.currentUser) === null || _a === void 0 ? void 0 : _a.spfx) || this.context.pageContext.user;
                        if (!currentUser) {
                            return [2 /*return*/, false];
                        }
                        restUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('").concat(this.READ_ACTIONS_LIST, "')/items") +
                            "?$filter=MessageId eq ".concat(messageId, " and UserEmail eq '").concat(currentUser.email, "'") +
                            "&$top=1";
                        return [4 /*yield*/, this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _b.sent();
                        if (!response.ok) {
                            return [2 /*return*/, false];
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _b.sent();
                        return [2 /*return*/, data.value.length > 0];
                    case 3:
                        error_13 = _b.sent();
                        console.error("EnhancedDataService: Error checking read status for message ".concat(messageId, ":"), error_13);
                        return [2 /*return*/, false];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get current user information for display
     */
    EnhancedDataService.prototype.getCurrentUser = function () {
        return this.currentUser;
    };
    /**
     * Create a new message (admin function)
     */
    EnhancedDataService.prototype.createMessage = function (message) {
        return __awaiter(this, void 0, void 0, function () {
            var newMessage, restUrl, spOpts, _a, response, errorDetails, errorBody, e_1, data, error_14, errorMsg;
            var _b, _c;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _d.trys.push([0, 9, , 10]);
                        newMessage = {
                            Title: message.Title,
                        };
                        // Add optional fields only if they have values
                        if (message.MessageContent) {
                            newMessage.MessageContent = message.MessageContent;
                        }
                        if (message.Priority) {
                            newMessage.Priority = message.Priority;
                        }
                        if (message.TargetAudience) {
                            newMessage.TargetAudience = message.TargetAudience;
                        }
                        // Add Source field if provided
                        if (message.Source) {
                            newMessage.Source = message.Source;
                        }
                        console.log('EnhancedDataService: Creating message with data:', newMessage);
                        console.log('EnhancedDataService: Target site URL:', this.getCurrentSiteUrl());
                        restUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('Important Messages')/items");
                        console.log('EnhancedDataService: REST API URL:', restUrl);
                        _b = {};
                        _c = {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-type': 'application/json;odata=nometadata',
                            'odata-version': ''
                        };
                        _a = 'X-RequestDigest';
                        return [4 /*yield*/, this.getRequestDigest()];
                    case 1:
                        spOpts = (_b.headers = (_c[_a] = _d.sent(),
                            _c),
                            _b.body = JSON.stringify(newMessage),
                            _b);
                        return [4 /*yield*/, this.context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, spOpts)];
                    case 2:
                        response = _d.sent();
                        if (!!response.ok) return [3 /*break*/, 7];
                        console.error('SharePoint API Request Failed:');
                        console.error('- Status:', response.status, response.statusText);
                        console.error('- URL:', restUrl);
                        console.error('- Data sent:', newMessage);
                        errorDetails = "".concat(response.status, " ").concat(response.statusText);
                        _d.label = 3;
                    case 3:
                        _d.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, response.text()];
                    case 4:
                        errorBody = _d.sent();
                        console.error('SharePoint API Error Details:', errorBody);
                        errorDetails += " - ".concat(errorBody);
                        return [3 /*break*/, 6];
                    case 5:
                        e_1 = _d.sent();
                        console.error('Could not read error response body:', e_1);
                        return [3 /*break*/, 6];
                    case 6: throw new Error("Failed to create message: ".concat(errorDetails));
                    case 7: return [4 /*yield*/, response.json()];
                    case 8:
                        data = _d.sent();
                        console.log("EnhancedDataService: Created new message with ID: ".concat(data.Id));
                        return [2 /*return*/, data.Id];
                    case 9:
                        error_14 = _d.sent();
                        console.error('EnhancedDataService: Error creating message:', error_14);
                        errorMsg = error_14 instanceof Error ? error_14.message : 'Failed to create message';
                        throw new Error(errorMsg);
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get a specific message by ID
     */
    EnhancedDataService.prototype.getMessageById = function (messageId) {
        return __awaiter(this, void 0, void 0, function () {
            var restUrl, response, data, error_15;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        restUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/items(").concat(messageId, ")") +
                            "?$select=Id,Title,MessageContent,Priority,TargetAudience,ReadBy,Created,Modified";
                        return [4 /*yield*/, this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP ".concat(response.status, ": ").concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, this.mapToMessage(data)];
                    case 3:
                        error_15 = _a.sent();
                        console.error("EnhancedDataService: Error fetching message ".concat(messageId, ":"), error_15);
                        throw new Error("Failed to fetch message with ID ".concat(messageId));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if user has access to specific functionality based on Graph groups
     */
    EnhancedDataService.prototype.hasUserRole = function (role) {
        var _a;
        if (!this.currentUser) {
            return false;
        }
        var userGroups = this.currentUser.groups || [];
        var userJobTitle = ((_a = this.currentUser.graph) === null || _a === void 0 ? void 0 : _a.jobTitle) || '';
        // Check if user is in specific groups or has specific job titles
        switch (role.toLowerCase()) {
            case 'admin':
            case 'administrator':
                return userGroups.some(function (group) {
                    return group.toLowerCase().includes('admin') ||
                        group.toLowerCase().includes('administrator');
                }) || userJobTitle.toLowerCase().includes('admin');
            case 'manager':
                return userGroups.some(function (group) {
                    return group.toLowerCase().includes('manager') ||
                        group.toLowerCase().includes('lead');
                }) || userJobTitle.toLowerCase().includes('manager');
            case 'hr':
                return userGroups.some(function (group) {
                    return group.toLowerCase().includes('hr') ||
                        group.toLowerCase().includes('human');
                }) || userJobTitle.toLowerCase().includes('hr');
            default:
                return userGroups.some(function (group) {
                    return group.toLowerCase().includes(role.toLowerCase());
                });
        }
    };
    // Include all the private helper methods from the original DataService
    EnhancedDataService.prototype.getMessagesWithProgressiveQuerying = function (siteUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var availableColumns, basicQuery, enhancedQuery, error_16;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log('EnhancedDataService: Starting progressive column querying...');
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 7, , 8]);
                        return [4 /*yield*/, this.getAvailableColumns(siteUrl)];
                    case 2:
                        availableColumns = _a.sent();
                        console.log('EnhancedDataService: Available columns detected:', availableColumns);
                        return [4 /*yield*/, this.testBasicQuery(siteUrl)];
                    case 3:
                        basicQuery = _a.sent();
                        if (!basicQuery.success) return [3 /*break*/, 5];
                        console.log('EnhancedDataService: Basic query successful, trying enhanced query...');
                        return [4 /*yield*/, this.testEnhancedQuery(siteUrl, availableColumns)];
                    case 4:
                        enhancedQuery = _a.sent();
                        if (enhancedQuery.success) {
                            return [2 /*return*/, enhancedQuery.data];
                        }
                        else {
                            console.log('EnhancedDataService: Enhanced query failed, using basic data');
                            return [2 /*return*/, basicQuery.data];
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        console.log('EnhancedDataService: Basic query failed, using mock data');
                        return [2 /*return*/, this.getMockMessages()];
                    case 6: return [3 /*break*/, 8];
                    case 7:
                        error_16 = _a.sent();
                        console.error('EnhancedDataService: Progressive querying failed:', error_16);
                        return [2 /*return*/, this.getMockMessages()];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    EnhancedDataService.prototype.testBasicQuery = function (siteUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var restUrl, response, data, messages, error_17;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        restUrl = "".concat(siteUrl, "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/items") +
                            "?$select=Id,Title,Created,Modified" +
                            "&$top=10" +
                            "&$orderby=Created desc";
                        return [4 /*yield*/, this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, { success: false, data: [] }];
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        messages = data.value.map(function (item) { return _this.mapToMessageBasic(item); });
                        return [2 /*return*/, { success: true, data: messages }];
                    case 3:
                        error_17 = _a.sent();
                        console.error('EnhancedDataService: Basic query error:', error_17);
                        return [2 /*return*/, { success: false, data: [] }];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    EnhancedDataService.prototype.testEnhancedQuery = function (siteUrl, availableColumns) {
        return __awaiter(this, void 0, void 0, function () {
            var safeColumns, restUrl, response, data, messages, error_18;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        safeColumns = this.buildConservativeColumnQuery(availableColumns);
                        restUrl = "".concat(siteUrl, "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/items") +
                            "?$select=".concat(safeColumns.select) +
                            (safeColumns.expand ? "&$expand=".concat(safeColumns.expand) : '') +
                            "&$top=50" +
                            "&$orderby=Created desc";
                        return [4 /*yield*/, this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, { success: false, data: [] }];
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        messages = data.value.map(function (item) { return _this.mapToMessageWithAvailableColumns(item, availableColumns); });
                        return [2 /*return*/, { success: true, data: messages }];
                    case 3:
                        error_18 = _a.sent();
                        console.error('EnhancedDataService: Enhanced query error:', error_18);
                        return [2 /*return*/, { success: false, data: [] }];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    EnhancedDataService.prototype.buildConservativeColumnQuery = function (availableColumns) {
        var safeColumns = ['Id', 'Title', 'Created', 'Modified'];
        var criticalFields = ['MessageContent', 'Priority', 'TargetAudience', 'ReadBy'];
        for (var _i = 0, criticalFields_1 = criticalFields; _i < criticalFields_1.length; _i++) {
            var field = criticalFields_1[_i];
            if (availableColumns.indexOf(field) !== -1) {
                safeColumns.push(field);
            }
        }
        var contentFields = ['Body', 'Description'];
        for (var _a = 0, contentFields_1 = contentFields; _a < contentFields_1.length; _a++) {
            var field = contentFields_1[_a];
            if (availableColumns.indexOf(field) !== -1 && safeColumns.indexOf('MessageContent') === -1) {
                safeColumns.push(field);
                break;
            }
        }
        var expandAuthor = false;
        if (availableColumns.indexOf('Author') !== -1) {
            safeColumns.push('Author/Title');
            expandAuthor = true;
        }
        return {
            select: safeColumns.join(','),
            expand: expandAuthor ? 'Author' : ''
        };
    };
    EnhancedDataService.prototype.getAvailableColumns = function (siteUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var listInfoUrl, response, listInfo, columns, error_19;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        listInfoUrl = "".concat(siteUrl, "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false");
                        return [4 /*yield*/, this.context.spHttpClient.get(listInfoUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, response.json()];
                    case 2:
                        listInfo = _a.sent();
                        columns = listInfo.value.map(function (field) { return field.InternalName; });
                        return [2 /*return*/, columns];
                    case 3: return [2 /*return*/, ['Id', 'Title', 'Created', 'Modified']];
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        error_19 = _a.sent();
                        console.warn('EnhancedDataService: Error fetching column schema:', error_19);
                        return [2 /*return*/, ['Id', 'Title', 'Created', 'Modified']];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    EnhancedDataService.prototype.mapToMessageBasic = function (item) {
        return {
            Id: item.Id || 0,
            Title: item.Title || 'Untitled Message',
            MessageContent: 'Click to view message details',
            Priority: 'Medium',
            TargetAudience: 'All Users',
            ReadBy: '',
            ReadById: [],
            Created: new Date(item.Created),
            Modified: new Date(item.Modified),
            Author: {
                Title: 'Unknown',
                Email: ''
            }
        };
    };
    EnhancedDataService.prototype.mapToMessageWithAvailableColumns = function (item, availableColumns) {
        var _a, _b;
        var getFieldValue = function (fieldNames, defaultValue) {
            if (defaultValue === void 0) { defaultValue = null; }
            for (var _i = 0, fieldNames_1 = fieldNames; _i < fieldNames_1.length; _i++) {
                var fieldName = fieldNames_1[_i];
                if (item[fieldName] !== undefined && item[fieldName] !== null) {
                    return item[fieldName];
                }
            }
            return defaultValue;
        };
        return {
            Id: item.Id || 0,
            Title: item.Title || 'Untitled Message',
            MessageContent: getFieldValue(['MessageContent', 'Body', 'Description', 'Content'], 'No content available'),
            Priority: getFieldValue(['Priority', 'Importance', 'Level'], 'Medium'),
            TargetAudience: getFieldValue(['TargetAudience', 'Audience', 'Group'], 'All Users'),
            ReadBy: getFieldValue(['ReadBy', 'ReadStatus'], ''),
            ReadById: getFieldValue(['ReadBy', 'ReadStatus']) ?
                getFieldValue(['ReadBy', 'ReadStatus']).split(';').map(function (email) { return email.trim(); }).filter(function (email) { return email; }) : [],
            Created: new Date(item.Created),
            Modified: new Date(item.Modified),
            Author: {
                Title: ((_a = item.Author) === null || _a === void 0 ? void 0 : _a.Title) || 'Unknown',
                Email: ((_b = item.Author) === null || _b === void 0 ? void 0 : _b.Email) || ''
            }
        };
    };
    EnhancedDataService.prototype.getRequestDigest = function () {
        return __awaiter(this, void 0, void 0, function () {
            var restUrl, response, data, error_20;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context || !this.context.spHttpClient) {
                            throw new Error('EnhancedDataService context not initialized.');
                        }
                        restUrl = "".concat(this.getCurrentSiteUrl(), "/_api/contextinfo");
                        return [4 /*yield*/, this.context.spHttpClient.post(restUrl, SPHttpClient.configurations.v1, {})];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Failed to get request digest: ".concat(response.status, " ").concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, data.FormDigestValue];
                    case 3:
                        error_20 = _a.sent();
                        console.error('EnhancedDataService: Error getting request digest:', error_20);
                        throw error_20;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    EnhancedDataService.prototype.updateMessageReadBy = function (messageId, userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var getUrl, getResponse, currentData, currentReadBy, emails, updatedReadBy, updateUrl, spOpts, _a, error_21;
            var _b, _c;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _d.trys.push([0, 6, , 7]);
                        getUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/items(").concat(messageId, ")") +
                            "?$select=ReadBy";
                        return [4 /*yield*/, this.context.spHttpClient.get(getUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        getResponse = _d.sent();
                        if (!getResponse.ok) {
                            console.warn('Could not fetch current ReadBy value');
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, getResponse.json()];
                    case 2:
                        currentData = _d.sent();
                        currentReadBy = currentData.ReadBy || '';
                        emails = currentReadBy ? currentReadBy.split(';').filter(function (email) { return email.trim(); }) : [];
                        if (!!emails.includes(userEmail)) return [3 /*break*/, 5];
                        emails.push(userEmail);
                        updatedReadBy = emails.join(';');
                        updateUrl = "".concat(this.getCurrentSiteUrl(), "/_api/web/lists/getByTitle('").concat(this.MESSAGES_LIST, "')/items(").concat(messageId, ")");
                        _b = {};
                        _c = {
                            'Accept': 'application/json;odata=nometadata',
                            'Content-type': 'application/json;odata=nometadata',
                            'odata-version': '',
                            'IF-MATCH': '*',
                            'X-HTTP-Method': 'MERGE'
                        };
                        _a = 'X-RequestDigest';
                        return [4 /*yield*/, this.getRequestDigest()];
                    case 3:
                        spOpts = (_b.headers = (_c[_a] = _d.sent(),
                            _c),
                            _b.body = JSON.stringify({
                                ReadBy: updatedReadBy
                            }),
                            _b);
                        return [4 /*yield*/, this.context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, spOpts)];
                    case 4:
                        _d.sent();
                        _d.label = 5;
                    case 5: return [3 /*break*/, 7];
                    case 6:
                        error_21 = _d.sent();
                        console.error("Error updating ReadBy field for message ".concat(messageId, ":"), error_21);
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    EnhancedDataService.prototype.getDeviceInfo = function () {
        var userAgent = navigator.userAgent;
        var platform = navigator.platform;
        return "".concat(platform, " - ").concat(userAgent.substring(0, 100));
    };
    EnhancedDataService.prototype.getMockMessages = function () {
        var tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        var nextWeek = new Date();
        nextWeek.setDate(nextWeek.getDate() + 7);
        return [
            {
                Id: 1,
                Title: "🚀 New Feature Release - Enhanced with Graph Integration",
                MessageContent: "We're excited to announce the release of our new dashboard features with Microsoft Graph integration! No admin approval required for basic features.",
                Priority: "High",
                TargetAudience: "All Users",
                ReadBy: "",
                Created: new Date(),
                Modified: new Date(),
                Author: {
                    Title: "System Administrator",
                    Email: "admin@company.com"
                }
            },
            {
                Id: 2,
                Title: "📅 Maintenance Window Scheduled",
                MessageContent: "Scheduled maintenance will occur this weekend from 2 AM to 6 AM. The system will be temporarily unavailable during this time.",
                Priority: "Medium",
                TargetAudience: "All Users",
                ReadBy: "",
                Created: new Date(),
                Modified: new Date(),
                Author: {
                    Title: "IT Team",
                    Email: "it@company.com"
                }
            },
            {
                Id: 3,
                Title: "📊 Dashboard Tutorial Available - Graph Enhanced",
                MessageContent: "New to the dashboard? Check out our comprehensive tutorial featuring Microsoft Graph integration for enhanced user experience without requiring admin permissions.",
                Priority: "Low",
                TargetAudience: "New Users",
                ReadBy: "",
                Created: new Date(),
                Modified: new Date(),
                Author: {
                    Title: "Training Team",
                    Email: "training@company.com"
                }
            }
        ];
    };
    return EnhancedDataService;
}());
export { EnhancedDataService };
// Export singleton instance
export var enhancedDataService = new EnhancedDataService();
//# sourceMappingURL=EnhancedDataService.js.map