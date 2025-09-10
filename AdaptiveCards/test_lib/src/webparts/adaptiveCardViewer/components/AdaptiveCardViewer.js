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
import { Stack, PrimaryButton, DefaultButton, Spinner, SpinnerSize, Text, Icon } from '@fluentui/react';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { EmployeeMessageList } from './EmployeeMessageList';
import { ManagerDashboard } from './ManagerDashboard';
import { TeamsMessageCreator } from './TeamsMessageCreator';
import styles from './AdaptiveCardViewer.module.scss';
/**
 * Adaptive Card Viewer Component
 * Displays adaptive cards based on user role and selected view
 */
var AdaptiveCardViewer = /** @class */ (function (_super) {
    __extends(AdaptiveCardViewer, _super);
    function AdaptiveCardViewer(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            isLoading: true,
            currentView: 'messages',
            userRole: 'employee',
            error: undefined
        };
        return _this;
        // Enhanced services are initialized globally
    }
    AdaptiveCardViewer.prototype.componentDidMount = function () {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var currentUser, userRole, isManager, error_1, error_2;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        _d.trys.push([0, 6, , 7]);
                        console.log('🔄 AdaptiveCardViewer: Component mounting...');
                        // Initialize enhanced data service
                        return [4 /*yield*/, enhancedDataService.initialize(this.props.context, (_c = (_b = (_a = this.props.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.web) === null || _c === void 0 ? void 0 : _c.absoluteUrl)];
                    case 1:
                        // Initialize enhanced data service
                        _d.sent();
                        currentUser = enhancedDataService.getCurrentUser();
                        userRole = 'employee';
                        _d.label = 2;
                    case 2:
                        _d.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, enhancedDataService.isCurrentUserManager()];
                    case 3:
                        isManager = _d.sent();
                        userRole = isManager ? 'manager' : 'employee';
                        console.log("\uD83C\uDFAF Manager status from SharePoint list: ".concat(isManager));
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _d.sent();
                        console.warn('⚠️ Could not check manager status from SharePoint list:', error_1);
                        // Fallback to employee role if we can't check the list
                        userRole = 'employee';
                        return [3 /*break*/, 5];
                    case 5:
                        this.setState({
                            userRole: userRole,
                            isLoading: false,
                            error: undefined
                        });
                        console.log("\uD83C\uDFAF User role determined: ".concat(userRole));
                        return [3 /*break*/, 7];
                    case 6:
                        error_2 = _d.sent();
                        console.error('❌ Error during component mount:', error_2);
                        this.setState({
                            error: "Kunde inte ladda komponenten: ".concat(error_2.message),
                            isLoading: false,
                            userRole: 'employee' // Safe fallback
                        });
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    AdaptiveCardViewer.prototype.render = function () {
        var _this = this;
        var _a = this.state, isLoading = _a.isLoading, currentView = _a.currentView, userRole = _a.userRole, error = _a.error;
        if (isLoading) {
            return (React.createElement(Stack, { horizontalAlign: "center", verticalAlign: "center", style: { minHeight: '200px' } },
                React.createElement(Spinner, { size: SpinnerSize.large, label: "Laddar..." })));
        }
        if (error) {
            return (React.createElement("div", { className: styles.adaptiveCardViewer },
                React.createElement(Stack, { horizontalAlign: "center", verticalAlign: "center", style: { minHeight: '200px' } },
                    React.createElement(Icon, { iconName: "Error", style: { fontSize: '48px', color: '#d13438', marginBottom: '16px' } }),
                    React.createElement(Text, { variant: "large", style: { color: '#d13438', marginBottom: '8px' } }, "Ett fel uppstod"),
                    React.createElement(Text, { variant: "medium", style: { color: '#605e5c', textAlign: 'center' } }, error),
                    React.createElement(PrimaryButton, { text: "F\u00F6rs\u00F6k igen", iconProps: { iconName: 'Refresh' }, onClick: function () { return window.location.reload(); }, style: { marginTop: '16px' } }))));
        }
        return (React.createElement("div", { className: styles.adaptiveCardViewer },
            React.createElement(Stack, { style: { marginBottom: 10 } },
                React.createElement(Text, { variant: "small", style: { color: '#666' } },
                    "Inloggad som: ",
                    userRole === 'manager' ? '👑 Chef' : '👤 Medarbetare',
                    " (",
                    this.props.context.pageContext.user.email,
                    ")")),
            userRole === 'manager' ? (React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 16 }, style: { marginBottom: 20 } },
                React.createElement(PrimaryButton, { text: "\uD83D\uDCCA Manager Dashboard", onClick: function () { return _this.setState({ currentView: 'dashboard', error: undefined }); }, disabled: currentView === 'dashboard' }),
                React.createElement(DefaultButton, { text: "\u2709\uFE0F Skapa meddelande", onClick: function () { return _this.setState({ currentView: 'creator', error: undefined }); }, disabled: currentView === 'creator' }),
                React.createElement(DefaultButton, { text: "\uD83D\uDCEC Mina meddelanden", onClick: function () { return _this.setState({ currentView: 'messages', error: undefined }); }, disabled: currentView === 'messages' }),
                React.createElement(DefaultButton, { text: "\uD83D\uDD0D Diagnostik", onClick: function () { return _this.setState({ currentView: 'diagnostic', error: undefined }); }, disabled: currentView === 'diagnostic' }))) : (React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 16 }, style: { marginBottom: 20 } },
                React.createElement(PrimaryButton, { text: "\uD83D\uDCEC Mina meddelanden", onClick: function () { return _this.setState({ currentView: 'messages', error: undefined }); }, disabled: currentView === 'messages' }))),
            this.renderCurrentView()));
    };
    AdaptiveCardViewer.prototype.renderCurrentView = function () {
        var _this = this;
        var _a = this.state, currentView = _a.currentView, userRole = _a.userRole;
        try {
            switch (currentView) {
                case 'messages':
                    return (React.createElement(EmployeeMessageList, null));
                case 'dashboard':
                    if (userRole !== 'manager') {
                        return (React.createElement(Stack, { horizontalAlign: "center", style: { padding: '20px' } },
                            React.createElement(Text, { variant: "large", style: { color: '#d13438' } }, "\u274C Du har inte beh\u00F6righet att se denna sida")));
                    }
                    return (React.createElement(ManagerDashboard, null));
                case 'creator':
                    if (userRole !== 'manager') {
                        return (React.createElement(Stack, { horizontalAlign: "center", style: { padding: '20px' } },
                            React.createElement(Text, { variant: "large", style: { color: '#d13438' } }, "\u274C Du har inte beh\u00F6righet att skapa meddelanden")));
                    }
                    return (React.createElement(TeamsMessageCreator, { context: this.props.context, onMessageCreated: function () {
                            console.log('Message created, refreshing dashboard...');
                            _this.setState({ currentView: 'dashboard' });
                        } }));
                case 'diagnostic':
                    if (userRole !== 'manager') {
                        return (React.createElement(Stack, { horizontalAlign: "center", style: { padding: '20px' } },
                            React.createElement(Text, { variant: "large", style: { color: '#d13438' } }, "\u274C Du har inte beh\u00F6righet att se diagnostik")));
                    }
                    return (React.createElement(Stack, null,
                        React.createElement(Text, { variant: "large" }, "Diagnostic tools have been replaced with enhanced services."),
                        React.createElement(Text, null, "Use the Messages or Manager Dashboard views instead.")));
                default:
                    return (React.createElement(Stack, { horizontalAlign: "center", style: { padding: '20px' } },
                        React.createElement(Text, { variant: "large", style: { color: '#d13438' } },
                            "\u274C Ok\u00E4nd vy: ",
                            currentView)));
            }
        }
        catch (error) {
            console.error('❌ Error rendering view:', error);
            return (React.createElement(Stack, { horizontalAlign: "center", style: { padding: '20px' } },
                React.createElement(Text, { variant: "large", style: { color: '#d13438' } }, "\u274C Ett fel uppstod n\u00E4r sidan skulle laddas"),
                React.createElement(Text, { variant: "medium", style: { color: '#666' } }, error.message)));
        }
    };
    return AdaptiveCardViewer;
}(React.Component));
export default AdaptiveCardViewer;
//# sourceMappingURL=AdaptiveCardViewer.js.map