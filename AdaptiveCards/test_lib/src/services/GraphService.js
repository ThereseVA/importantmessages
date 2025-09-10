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
var GraphService = /** @class */ (function () {
    function GraphService() {
        this.graphClient = null;
    }
    /**
     * Initialize the Graph service with SPFx context
     */
    GraphService.prototype.initialize = function (context) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                this.context = context;
                try {
                    console.log('GraphService: Initializing with SPFx context');
                    // We'll use MSGraphClientV3 directly instead of creating a custom client
                    // This avoids authentication issues and uses SPFx's built-in Graph access
                    console.log('GraphService: Initialized successfully');
                }
                catch (error) {
                    console.error('GraphService: Failed to initialize:', error);
                    throw error;
                }
                return [2 /*return*/];
            });
        });
    };
    /**
     * Get current user information without requiring admin approval
     */
    GraphService.prototype.getCurrentUser = function () {
        var _a, _b;
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, user, error_1, contextUser;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        _c.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            console.warn('GraphService: Context not initialized');
                            return [2 /*return*/, null];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _c.sent();
                        return [4 /*yield*/, msGraphClient.api('/me').get()];
                    case 2:
                        user = _c.sent();
                        return [2 /*return*/, {
                                id: user.id || '',
                                displayName: user.displayName || '',
                                mail: user.mail || user.userPrincipalName || '',
                                userPrincipalName: user.userPrincipalName || '',
                                jobTitle: user.jobTitle,
                                department: user.department
                            }];
                    case 3:
                        error_1 = _c.sent();
                        console.warn('GraphService: Error getting current user, falling back to SPFx context:', error_1);
                        // Fallback to SPFx context user info
                        if ((_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.user) {
                            contextUser = this.context.pageContext.user;
                            return [2 /*return*/, {
                                    id: contextUser.loginName,
                                    displayName: contextUser.displayName,
                                    mail: contextUser.email,
                                    userPrincipalName: contextUser.loginName,
                                    jobTitle: undefined,
                                    department: undefined
                                }];
                        }
                        return [2 /*return*/, null];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get user's groups - this uses delegated permissions (user's own access)
     */
    GraphService.prototype.getCurrentUserGroups = function () {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, response, groups, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            console.warn('GraphService: Context not initialized');
                            return [2 /*return*/, []];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        return [4 /*yield*/, msGraphClient.api('/me/memberOf').get()];
                    case 2:
                        response = _a.sent();
                        groups = response.value || [];
                        return [2 /*return*/, groups
                                .filter(function (group) { return group['@odata.type'] === '#microsoft.graph.group'; })
                                .map(function (group) { return group.displayName || group.mailNickname || group.id; })];
                    case 3:
                        error_2 = _a.sent();
                        console.warn('GraphService: Error getting user groups, returning empty array:', error_2);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get SharePoint sites the user has access to - using delegated permissions
     */
    GraphService.prototype.getUserAccessibleSites = function () {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, response, sites, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            console.warn('GraphService: Context not initialized');
                            return [2 /*return*/, []];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        return [4 /*yield*/, msGraphClient.api('/me/followedSites').get()];
                    case 2:
                        response = _a.sent();
                        sites = response.value || [];
                        return [2 /*return*/, sites.map(function (site) { return ({
                                id: site.id || '',
                                displayName: site.displayName || '',
                                webUrl: site.webUrl || '',
                                siteCollection: site.siteCollection
                            }); })];
                    case 3:
                        error_3 = _a.sent();
                        console.error('GraphService: Error getting user sites:', error_3);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Send Teams message using delegated permissions
     */
    GraphService.prototype.sendTeamsMessage = function (chatId, message) {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            console.warn('GraphService: Context not initialized');
                            return [2 /*return*/, false];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        // Send a message to a Teams chat - requires user to be a member
                        return [4 /*yield*/, msGraphClient.api("/chats/".concat(chatId, "/messages")).post({
                                body: {
                                    content: message,
                                    contentType: 'text'
                                }
                            })];
                    case 2:
                        // Send a message to a Teams chat - requires user to be a member
                        _a.sent();
                        return [2 /*return*/, true];
                    case 3:
                        error_4 = _a.sent();
                        console.error('GraphService: Error sending Teams message:', error_4);
                        return [2 /*return*/, false];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if user has access to a specific SharePoint site
     */
    GraphService.prototype.checkSiteAccess = function (siteUrl) {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, url, hostname, sitePath, site, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            return [2 /*return*/, false];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        url = new URL(siteUrl);
                        hostname = url.hostname;
                        sitePath = url.pathname;
                        return [4 /*yield*/, msGraphClient.api("/sites/".concat(hostname, ":").concat(sitePath)).get()];
                    case 2:
                        site = _a.sent();
                        return [2 /*return*/, !!site];
                    case 3:
                        error_5 = _a.sent();
                        console.warn('GraphService: User does not have access to site:', siteUrl);
                        return [2 /*return*/, false];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get user photo as base64 string
     */
    GraphService.prototype.getUserPhoto = function () {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, photoResponse, arrayBuffer, uint8Array, base64, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        if (!this.context) {
                            return [2 /*return*/, null];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        return [4 /*yield*/, msGraphClient.api('/me/photo/$value').get()];
                    case 2:
                        photoResponse = _a.sent();
                        return [4 /*yield*/, photoResponse.arrayBuffer()];
                    case 3:
                        arrayBuffer = _a.sent();
                        uint8Array = new Uint8Array(arrayBuffer);
                        base64 = btoa(String.fromCharCode.apply(null, Array.from(uint8Array)));
                        return [2 /*return*/, "data:image/jpeg;base64,".concat(base64)];
                    case 4:
                        error_6 = _a.sent();
                        console.warn('GraphService: Could not get user photo:', error_6);
                        return [2 /*return*/, null];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Create a SharePoint list item using Graph API (if user has permissions)
     */
    GraphService.prototype.createListItem = function (siteId, listId, fields) {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, response, error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            throw new Error('GraphService not initialized');
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        return [4 /*yield*/, msGraphClient.api("/sites/".concat(siteId, "/lists/").concat(listId, "/items")).post({
                                fields: fields
                            })];
                    case 2:
                        response = _a.sent();
                        return [2 /*return*/, response];
                    case 3:
                        error_7 = _a.sent();
                        console.error('GraphService: Error creating list item:', error_7);
                        throw error_7;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get SharePoint list items using Graph API (if user has permissions)
     */
    GraphService.prototype.getListItems = function (siteId, listId, filter) {
        return __awaiter(this, void 0, void 0, function () {
            var msGraphClientFactory, msGraphClient, apiCall, response, error_8;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        if (!this.context) {
                            return [2 /*return*/, []];
                        }
                        msGraphClientFactory = this.context.msGraphClientFactory;
                        return [4 /*yield*/, msGraphClientFactory.getClient('3')];
                    case 1:
                        msGraphClient = _a.sent();
                        apiCall = msGraphClient.api("/sites/".concat(siteId, "/lists/").concat(listId, "/items")).expand('fields');
                        if (filter) {
                            apiCall = apiCall.filter(filter);
                        }
                        return [4 /*yield*/, apiCall.get()];
                    case 2:
                        response = _a.sent();
                        return [2 /*return*/, response.value || []];
                    case 3:
                        error_8 = _a.sent();
                        console.error('GraphService: Error getting list items:', error_8);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if the service is properly initialized
     */
    GraphService.prototype.isInitialized = function () {
        return !!this.context;
    };
    /**
     * Get enhanced user information by combining Graph and SPFx context
     */
    GraphService.prototype.getEnhancedUserInfo = function () {
        var _a, _b;
        return __awaiter(this, void 0, void 0, function () {
            var _c, graphUser, groups, photo, isManager, isAdmin;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0: return [4 /*yield*/, Promise.all([
                            this.getCurrentUser(),
                            this.getCurrentUserGroups(),
                            this.getUserPhoto().catch(function () { return null; })
                        ])];
                    case 1:
                        _c = _d.sent(), graphUser = _c[0], groups = _c[1], photo = _c[2];
                        isManager = this.checkManagerStatus(graphUser, groups);
                        isAdmin = this.checkAdminStatus(groups);
                        return [2 /*return*/, {
                                graph: graphUser,
                                context: ((_b = (_a = this.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.user) ? {
                                    displayName: this.context.pageContext.user.displayName,
                                    email: this.context.pageContext.user.email,
                                    loginName: this.context.pageContext.user.loginName
                                } : null,
                                groups: groups,
                                photo: photo || undefined,
                                isManager: isManager,
                                isAdmin: isAdmin
                            }];
                }
            });
        });
    };
    /**
     * Check if user has manager status based on job title or groups
     */
    GraphService.prototype.checkManagerStatus = function (user, groups) {
        var _a;
        if (!user)
            return false;
        // Check job title for manager keywords
        var managerTitles = ['manager', 'director', 'supervisor', 'lead', 'head'];
        var jobTitle = ((_a = user.jobTitle) === null || _a === void 0 ? void 0 : _a.toLowerCase()) || '';
        if (managerTitles.some(function (title) { return jobTitle.includes(title); })) {
            return true;
        }
        // Check groups for manager groups (customize as needed)
        var managerGroups = ['managers', 'leadership', 'directors'];
        return groups.some(function (group) {
            return managerGroups.some(function (managerGroup) {
                return group.toLowerCase().includes(managerGroup);
            });
        });
    };
    /**
     * Check if user has admin status based on groups
     */
    GraphService.prototype.checkAdminStatus = function (groups) {
        var adminGroups = ['administrators', 'admin', 'global administrators'];
        return groups.some(function (group) {
            return adminGroups.some(function (adminGroup) {
                return group.toLowerCase().includes(adminGroup);
            });
        });
    };
    return GraphService;
}());
export { GraphService };
// Export singleton instance
export var graphService = new GraphService();
//# sourceMappingURL=GraphService.js.map