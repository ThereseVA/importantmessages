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
var ManagersListService = /** @class */ (function () {
    function ManagersListService(context) {
        this.listName = "Managers";
        this.context = context;
    }
    /**
     * Get all active managers from the Managers list
     */
    ManagersListService.prototype.getActiveManagers = function () {
        return __awaiter(this, void 0, void 0, function () {
            var siteUrl, listUrl, response, data, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        siteUrl = this.context.pageContext.web.absoluteUrl;
                        listUrl = "".concat(siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items?$select=Id,Title,ManagersEmail/EMail,ManagersEmail/Title,ManagersDisplayName,Department,ManagerLevel,IsActive,StartDate,EndDate,Notes&$expand=ManagersEmail&$filter=IsActive eq true&$orderby=ManagersDisplayName");
                        return [4 /*yield*/, this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP error! status: ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, data.value];
                    case 3:
                        error_1 = _a.sent();
                        console.error('Error fetching managers from SharePoint list:', error_1);
                        throw error_1;
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Check if a specific user is a manager
     */
    ManagersListService.prototype.isUserManager = function (userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var managers, isManager, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getActiveManagers()];
                    case 1:
                        managers = _a.sent();
                        // Enhanced debugging
                        console.log('🔍 Manager Access Debug - Checking user email:', userEmail);
                        console.log('🔍 Total active managers found:', managers.length);
                        managers.forEach(function (manager, index) {
                            var _a, _b;
                            console.log("\uD83D\uDD0D Manager ".concat(index + 1, ":"), {
                                displayName: manager.ManagersDisplayName,
                                managerEmailObject: manager.ManagersEmail,
                                emailProperty: (_a = manager.ManagersEmail) === null || _a === void 0 ? void 0 : _a.EMail,
                                titleProperty: (_b = manager.ManagersEmail) === null || _b === void 0 ? void 0 : _b.Title,
                                fullObject: JSON.stringify(manager.ManagersEmail, null, 2)
                            });
                        });
                        isManager = managers.some(function (manager) {
                            var _a, _b, _c, _d;
                            var emailFromEMail = (_b = (_a = manager.ManagersEmail) === null || _a === void 0 ? void 0 : _a.EMail) === null || _b === void 0 ? void 0 : _b.toLowerCase();
                            var emailFromTitle = (_d = (_c = manager.ManagersEmail) === null || _c === void 0 ? void 0 : _c.Title) === null || _d === void 0 ? void 0 : _d.toLowerCase();
                            var userEmailLower = userEmail.toLowerCase();
                            console.log("\uD83D\uDD0D Checking manager ".concat(manager.ManagersDisplayName, ":"), {
                                emailFromEMail: emailFromEMail,
                                emailFromTitle: emailFromTitle,
                                userEmailLower: userEmailLower,
                                matchesEMail: emailFromEMail === userEmailLower,
                                matchesTitle: emailFromTitle === userEmailLower
                            });
                            return emailFromEMail === userEmailLower || emailFromTitle === userEmailLower;
                        });
                        console.log('🔍 Final result - Is user a manager?', isManager);
                        return [2 /*return*/, isManager];
                    case 2:
                        error_2 = _a.sent();
                        console.error('Error checking if user is manager:', error_2);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get manager details for a specific user
     */
    ManagersListService.prototype.getManagerDetails = function (userEmail) {
        return __awaiter(this, void 0, void 0, function () {
            var managers, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getActiveManagers()];
                    case 1:
                        managers = _a.sent();
                        return [2 /*return*/, managers.find(function (manager) { var _a, _b; return ((_b = (_a = manager.ManagersEmail) === null || _a === void 0 ? void 0 : _a.EMail) === null || _b === void 0 ? void 0 : _b.toLowerCase()) === userEmail.toLowerCase(); }) || null];
                    case 2:
                        error_3 = _a.sent();
                        console.error('Error getting manager details:', error_3);
                        return [2 /*return*/, null];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get managers by department
     */
    ManagersListService.prototype.getManagersByDepartment = function (department) {
        return __awaiter(this, void 0, void 0, function () {
            var managers, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getActiveManagers()];
                    case 1:
                        managers = _a.sent();
                        return [2 /*return*/, managers.filter(function (manager) { var _a; return ((_a = manager.Department) === null || _a === void 0 ? void 0 : _a.toLowerCase()) === department.toLowerCase(); })];
                    case 2:
                        error_4 = _a.sent();
                        console.error('Error getting managers by department:', error_4);
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Get managers by level
     */
    ManagersListService.prototype.getManagersByLevel = function (level) {
        return __awaiter(this, void 0, void 0, function () {
            var managers, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.getActiveManagers()];
                    case 1:
                        managers = _a.sent();
                        return [2 /*return*/, managers.filter(function (manager) { var _a; return ((_a = manager.ManagerLevel) === null || _a === void 0 ? void 0 : _a.toLowerCase()) === level.toLowerCase(); })];
                    case 2:
                        error_5 = _a.sent();
                        console.error('Error getting managers by level:', error_5);
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Add a new manager to the list
     */
    ManagersListService.prototype.addManager = function (manager) {
        return __awaiter(this, void 0, void 0, function () {
            var siteUrl, listUrl, body, response, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        siteUrl = this.context.pageContext.web.absoluteUrl;
                        listUrl = "".concat(siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items");
                        body = JSON.stringify({
                            Title: manager.Title || manager.ManagersDisplayName,
                            ManagersDisplayName: manager.ManagersDisplayName,
                            Department: manager.Department,
                            ManagerLevel: manager.ManagerLevel,
                            IsActive: manager.IsActive !== undefined ? manager.IsActive : true,
                            StartDate: manager.StartDate,
                            EndDate: manager.EndDate,
                            Notes: manager.Notes
                        });
                        return [4 /*yield*/, this.context.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=verbose',
                                    'odata-version': ''
                                },
                                body: body
                            })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, response.ok];
                    case 2:
                        error_6 = _a.sent();
                        console.error('Error adding manager to SharePoint list:', error_6);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Update an existing manager
     */
    ManagersListService.prototype.updateManager = function (managerId, updates) {
        return __awaiter(this, void 0, void 0, function () {
            var siteUrl, listUrl, getResponse, etag, body, response, error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        siteUrl = this.context.pageContext.web.absoluteUrl;
                        listUrl = "".concat(siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/items(").concat(managerId, ")");
                        return [4 /*yield*/, this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        getResponse = _a.sent();
                        if (!getResponse.ok) {
                            throw new Error('Manager not found');
                        }
                        etag = getResponse.headers.get('ETag');
                        body = JSON.stringify({
                            Title: updates.Title,
                            ManagersDisplayName: updates.ManagersDisplayName,
                            Department: updates.Department,
                            ManagerLevel: updates.ManagerLevel,
                            IsActive: updates.IsActive,
                            StartDate: updates.StartDate,
                            EndDate: updates.EndDate,
                            Notes: updates.Notes
                        });
                        return [4 /*yield*/, this.context.spHttpClient.post(listUrl, SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json;odata=nometadata',
                                    'Content-type': 'application/json;odata=verbose',
                                    'odata-version': '',
                                    'X-HTTP-Method': 'MERGE',
                                    'If-Match': etag || '*'
                                },
                                body: body
                            })];
                    case 2:
                        response = _a.sent();
                        return [2 /*return*/, response.ok];
                    case 3:
                        error_7 = _a.sent();
                        console.error('Error updating manager in SharePoint list:', error_7);
                        return [2 /*return*/, false];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Deactivate a manager (set IsActive to false)
     */
    ManagersListService.prototype.deactivateManager = function (managerId, endDate) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.updateManager(managerId, {
                            IsActive: false,
                            EndDate: endDate || new Date().toISOString()
                        })];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    /**
     * Check if the current user has permission to manage the Managers list
     */
    ManagersListService.prototype.canManageManagersList = function () {
        return __awaiter(this, void 0, void 0, function () {
            var siteUrl, listUrl, response, data, error_8;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        siteUrl = this.context.pageContext.web.absoluteUrl;
                        listUrl = "".concat(siteUrl, "/_api/web/lists/getbytitle('").concat(this.listName, "')/effectivebasepermissions");
                        return [4 /*yield*/, this.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, false];
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        // Check for AddListItems permission (value 2)
                        return [2 /*return*/, (data.High & 0) !== 0 || (data.Low & 2) !== 0];
                    case 3:
                        error_8 = _a.sent();
                        console.error('Error checking permissions for Managers list:', error_8);
                        return [2 /*return*/, false];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return ManagersListService;
}());
export { ManagersListService };
//# sourceMappingURL=ManagersListService.js.map