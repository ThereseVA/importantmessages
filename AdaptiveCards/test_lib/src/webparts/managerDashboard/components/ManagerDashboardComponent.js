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
import { enhancedDataService } from '../../../services/EnhancedDataService';
var ManagerDashboardComponent = /** @class */ (function (_super) {
    __extends(ManagerDashboardComponent, _super);
    function ManagerDashboardComponent(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            isManager: false,
            loading: true,
            error: null
        };
        return _this;
    }
    ManagerDashboardComponent.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var isManager, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        if (!this.props.context) return [3 /*break*/, 2];
                        return [4 /*yield*/, enhancedDataService.initialize(this.props.context)];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [4 /*yield*/, enhancedDataService.isCurrentUserManager()];
                    case 3:
                        isManager = _a.sent();
                        console.log('ManagerDashboard: Manager status from SharePoint list:', isManager);
                        this.setState({
                            isManager: isManager,
                            loading: false,
                            error: null
                        });
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _a.sent();
                        console.error('Error checking manager status from SharePoint list:', error_1);
                        this.setState({
                            isManager: false,
                            loading: false,
                            error: 'Failed to verify manager access from SharePoint Managers list'
                        });
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ManagerDashboardComponent.prototype.render = function () {
        var _a = this.state, loading = _a.loading, isManager = _a.isManager, error = _a.error;
        if (loading) {
            return (React.createElement("div", { style: { padding: '20px', textAlign: 'center' } },
                React.createElement("div", { style: { fontSize: '16px', color: '#666' } }, "\uD83D\uDD04 Checking permissions...")));
        }
        if (error) {
            return (React.createElement("div", { style: { padding: '20px', border: '1px solid #d13438', borderRadius: '4px', backgroundColor: '#fdf2f2' } },
                React.createElement("h3", { style: { color: '#d13438', margin: '0 0 10px 0' } }, "\u26A0\uFE0F Access Error"),
                React.createElement("p", { style: { margin: 0, color: '#666' } }, error)));
        }
        if (!isManager) {
            return (React.createElement("div", { style: { padding: '30px', textAlign: 'center', border: '1px solid #ffbe00', borderRadius: '8px', backgroundColor: '#fffbf0' } },
                React.createElement("div", { style: { fontSize: '48px', marginBottom: '16px' } }, "\uD83D\uDD12"),
                React.createElement("h3", { style: { color: '#d83b01', margin: '0 0 12px 0' } }, "Manager Access Required"),
                React.createElement("p", { style: { margin: '0 0 12px 0', color: '#666', fontSize: '14px' } }, "This Manager Dashboard is only accessible to managers listed in the SharePoint Managers list."),
                React.createElement("div", { style: {
                        padding: '12px',
                        backgroundColor: '#fff3cd',
                        borderRadius: '4px',
                        textAlign: 'left',
                        fontSize: '12px',
                        color: '#856404'
                    } },
                    React.createElement("strong", null, "How manager access is determined:"),
                    React.createElement("ul", { style: { margin: '8px 0 0 0', paddingLeft: '20px' } },
                        React.createElement("li", null, "Your email must be listed in the \"Managers\" SharePoint list"),
                        React.createElement("li", null, "Your entry must have \"Is Active\" set to \"Yes\""),
                        React.createElement("li", null, "Contact HR or IT to be added to the managers list"))),
                React.createElement("p", { style: { margin: '12px 0 0 0', color: '#666', fontSize: '12px' } }, "Contact your administrator if you believe you should have manager access.")));
        }
        return (React.createElement("div", null,
            React.createElement("div", { style: {
                    marginBottom: '20px',
                    padding: '16px',
                    backgroundColor: '#f0f8ff',
                    border: '1px solid #107c10',
                    borderRadius: '4px'
                } },
                React.createElement("h2", { style: { margin: '0 0 8px 0', color: '#107c10' } }, "\uD83C\uDF9B\uFE0F Manager Dashboard"),
                React.createElement("p", { style: { margin: 0, color: '#666', fontSize: '14px' } }, "Comprehensive management tools for messages, analytics, and system administration.")),
            React.createElement("div", { style: { padding: '20px', border: '1px solid #ddd', borderRadius: '8px' } },
                React.createElement("h3", null, "\uD83D\uDCCA Management Features"),
                React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginTop: '16px' } },
                    React.createElement("div", { style: { padding: '16px', border: '1px solid #107c10', borderRadius: '4px', backgroundColor: '#f8f9fa' } },
                        React.createElement("h4", { style: { margin: '0 0 8px 0', color: '#107c10' } }, "\uD83D\uDCC8 Message Analytics"),
                        React.createElement("p", { style: { margin: 0, fontSize: '14px', color: '#666' } }, "View comprehensive statistics about message delivery and engagement.")),
                    React.createElement("div", { style: { padding: '16px', border: '1px solid #0078d4', borderRadius: '4px', backgroundColor: '#f8f9fa' } },
                        React.createElement("h4", { style: { margin: '0 0 8px 0', color: '#0078d4' } }, "\uD83D\uDCDD Message Creation"),
                        React.createElement("p", { style: { margin: 0, fontSize: '14px', color: '#666' } }, "Create and send new messages to specific teams or user groups.")),
                    React.createElement("div", { style: { padding: '16px', border: '1px solid #d83b01', borderRadius: '4px', backgroundColor: '#f8f9fa' } },
                        React.createElement("h4", { style: { margin: '0 0 8px 0', color: '#d83b01' } }, "\u2699\uFE0F System Settings"),
                        React.createElement("p", { style: { margin: 0, fontSize: '14px', color: '#666' } }, "Configure system-wide settings and manage user permissions.")),
                    React.createElement("div", { style: { padding: '16px', border: '1px solid #8b5cf6', borderRadius: '4px', backgroundColor: '#f8f9fa' } },
                        React.createElement("h4", { style: { margin: '0 0 8px 0', color: '#8b5cf6' } }, "\uD83D\uDC65 User Management"),
                        React.createElement("p", { style: { margin: 0, fontSize: '14px', color: '#666' } }, "Manage user roles, groups, and access permissions."))),
                React.createElement("div", { style: { marginTop: '24px', padding: '16px', backgroundColor: '#fff3cd', border: '1px solid #ffbe00', borderRadius: '4px' } },
                    React.createElement("h4", { style: { margin: '0 0 8px 0', color: '#856404' } }, "\uD83D\uDEA7 Advanced Features Coming Soon"),
                    React.createElement("p", { style: { margin: 0, fontSize: '14px', color: '#856404' } }, "More advanced management features are in development. This dashboard will be expanded with additional capabilities.")),
                React.createElement("div", { style: { marginTop: '16px', textAlign: 'center' } },
                    React.createElement("p", { style: { fontSize: '12px', color: '#666' } },
                        "For access to the full Manager Dashboard functionality, use the Adaptive Card Viewer with ",
                        React.createElement("code", null, "?cardSource=manager-dashboard"),
                        " URL parameter.")))));
    };
    return ManagerDashboardComponent;
}(React.Component));
export { ManagerDashboardComponent };
//# sourceMappingURL=ManagerDashboardComponent.js.map