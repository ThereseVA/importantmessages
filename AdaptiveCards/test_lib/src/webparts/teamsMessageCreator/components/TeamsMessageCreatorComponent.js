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
import { TeamsMessageCreator } from '../../adaptiveCardViewer/components/TeamsMessageCreator';
import { EnhancedDataService } from '../../../services/EnhancedDataService';
var TeamsMessageCreatorComponent = /** @class */ (function (_super) {
    __extends(TeamsMessageCreatorComponent, _super);
    function TeamsMessageCreatorComponent(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            isManager: false,
            loading: true,
            error: null
        };
        _this.dataService = new EnhancedDataService();
        return _this;
    }
    TeamsMessageCreatorComponent.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            var userInfo, isManager, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        return [4 /*yield*/, this.dataService.initialize(this.props.context)];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.dataService.getCurrentUser()];
                    case 2:
                        userInfo = _a.sent();
                        isManager = userInfo.isManager || userInfo.isAdmin;
                        console.log('👤 Teams Message Creator user check:', {
                            displayName: userInfo.displayName,
                            isManager: isManager,
                            isAdmin: userInfo.isAdmin
                        });
                        this.setState({
                            isManager: isManager,
                            loading: false
                        });
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        console.error('❌ Error checking user permissions:', error_1);
                        this.setState({
                            error: error_1.message || 'Failed to verify permissions',
                            loading: false
                        });
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    TeamsMessageCreatorComponent.prototype.render = function () {
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
                React.createElement("p", { style: { margin: 0, color: '#666', fontSize: '14px' } }, "This Teams Message Creator is only accessible to managers and administrators."),
                React.createElement("p", { style: { margin: '12px 0 0 0', color: '#666', fontSize: '12px' } }, "Contact your administrator if you need access to this feature.")));
        }
        return (React.createElement("div", null,
            React.createElement("div", { style: {
                    marginBottom: '20px',
                    padding: '16px',
                    backgroundColor: '#f3f9ff',
                    border: '1px solid #0078d4',
                    borderRadius: '4px'
                } },
                React.createElement("h2", { style: { margin: '0 0 8px 0', color: '#0078d4' } }, "\uD83D\uDCDD Teams Message Creator"),
                React.createElement("p", { style: { margin: 0, color: '#666', fontSize: '14px' } }, "Create and send messages to Teams channels and user groups.")),
            React.createElement(TeamsMessageCreator, { context: this.props.context, dataService: this.dataService })));
    };
    return TeamsMessageCreatorComponent;
}(React.Component));
export { TeamsMessageCreatorComponent };
//# sourceMappingURL=TeamsMessageCreatorComponent.js.map