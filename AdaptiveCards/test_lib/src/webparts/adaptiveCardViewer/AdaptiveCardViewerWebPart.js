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
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'AdaptiveCardViewerWebPartStrings';
import { AdaptiveCardComponent } from './components/AdaptiveCardComponent';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { enhancedDataService } from '../../services/EnhancedDataService';
var AdaptiveCardViewerWebPart = /** @class */ (function (_super) {
    __extends(AdaptiveCardViewerWebPart, _super);
    function AdaptiveCardViewerWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    AdaptiveCardViewerWebPart.prototype.onInit = function () {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var currentUser, error_1;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        console.log('🚀 AdaptiveCardViewerWebPart.onInit() - Initializing Enhanced Data Service');
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 3, , 4]);
                        // Initialize the enhanced data service with Graph integration
                        return [4 /*yield*/, enhancedDataService.initialize(this.context)];
                    case 2:
                        // Initialize the enhanced data service with Graph integration
                        _d.sent();
                        console.log('✅ Enhanced Data Service initialized successfully');
                        currentUser = enhancedDataService.getCurrentUser();
                        if (currentUser) {
                            console.log('👤 Current user initialized:', {
                                displayName: ((_a = currentUser.spfx) === null || _a === void 0 ? void 0 : _a.displayName) || ((_b = currentUser.graph) === null || _b === void 0 ? void 0 : _b.displayName),
                                groups: ((_c = currentUser.groups) === null || _c === void 0 ? void 0 : _c.length) || 0,
                                hasPhoto: currentUser.hasPhoto
                            });
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _d.sent();
                        console.error('❌ Error initializing Enhanced Data Service:', error_1);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/, _super.prototype.onInit.call(this)];
                }
            });
        });
    };
    AdaptiveCardViewerWebPart.prototype.render = function () {
        var _this = this;
        console.log('🚀 AdaptiveCardViewerWebPart.render() called');
        console.log('🔧 Properties:', this.properties);
        console.log('🌐 Context available:', !!this.context);
        console.log('📱 Display mode:', this.displayMode);
        console.log('🎯 DOM element:', this.domElement);
        // Determine card URL based on source selection
        var cardJsonUrl;
        // Check for URL parameter override first
        var urlParams = new URLSearchParams(window.location.search);
        var componentParam = urlParams.get('component');
        var cardSource = this.properties.cardSource || 'manager-dashboard';
        // Override cardSource with URL parameter if present
        if (componentParam) {
            switch (componentParam) {
                case 'teams-message-creator':
                case 'manager-dashboard':
                case 'message-list-diagnostic':
                    cardSource = componentParam;
                    console.log('📋 Card source overridden by URL parameter:', cardSource);
                    break;
                default:
                    console.log('📋 Unknown component parameter, using default:', cardSource);
            }
        }
        else {
            console.log('📋 Card source selected:', cardSource);
        }
        switch (cardSource) {
            case 'manager-dashboard':
                cardJsonUrl = 'component:manager-dashboard';
                break;
            case 'teams-message-creator':
                cardJsonUrl = 'component:teams-message-creator';
                break;
            case 'message-list-diagnostic':
                cardJsonUrl = 'component:message-list-diagnostic';
                break;
            case 'asset-sample':
                cardJsonUrl = 'asset:sample-card';
                break;
            case 'template-sample':
                cardJsonUrl = 'template:sample';
                break;
            case 'template-dashboard':
                cardJsonUrl = 'template:dashboard';
                break;
            case 'asset-project':
                cardJsonUrl = 'asset:project-status';
                break;
            case 'asset-team':
                cardJsonUrl = 'asset:team-notification';
                break;
            case 'asset-sales':
                cardJsonUrl = 'asset:sales-dashboard';
                break;
            case 'custom':
                cardJsonUrl = this.properties.cardJsonUrl || 'asset:sample-card';
                break;
            default:
                cardJsonUrl = 'component:manager-dashboard';
                break;
        }
        console.log('🔗 Final cardJsonUrl:', cardJsonUrl);
        try {
            var element = React.createElement(AdaptiveCardComponent, {
                cardJsonUrl: cardJsonUrl,
                title: this.properties.title || 'Adaptive Card Demo',
                enableActions: this.properties.enableActions !== false,
                context: this.context,
                displayMode: this.displayMode,
                updateProperty: function (value) {
                    _this.properties.title = value;
                }
            });
            console.log('⚛️ React element created successfully:', element);
            console.log('🎨 About to render to DOM element:', this.domElement);
            ReactDom.render(element, this.domElement);
            console.log('✅ ReactDom.render completed successfully');
        }
        catch (error) {
            console.error('❌ Error in render method:', error);
            // Render error message directly to DOM
            if (this.domElement) {
                this.domElement.innerHTML = "\n          <div style=\"color: red; padding: 20px; border: 1px solid red; margin: 10px;\">\n            <h3>Web Part Render Error</h3>\n            <p>Error: ".concat(error.message || error, "</p>\n            <p>Please check browser console for details.</p>\n          </div>\n        ");
            }
        }
    };
    AdaptiveCardViewerWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    AdaptiveCardViewerWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('title', {
                                    label: strings.TitleFieldLabel
                                }),
                                PropertyPaneDropdown('cardSource', {
                                    label: 'Select Tool',
                                    options: [
                                        { key: 'manager-dashboard', text: '🎛️ Manager Dashboard (CREATE MESSAGES)' },
                                        { key: 'teams-message-creator', text: '� Teams Message Creator' },
                                        { key: 'message-list-diagnostic', text: '🔍 Message List Diagnostic' },
                                        { key: 'asset-sample', text: '📋 Sample Card Demo' }
                                    ],
                                    selectedKey: this.properties.cardSource || 'manager-dashboard'
                                }),
                                PropertyPaneTextField('cardJsonUrl', {
                                    label: strings.CardJsonUrlFieldLabel,
                                    description: strings.CardJsonUrlFieldDescription,
                                    disabled: this.properties.cardSource !== 'custom'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return AdaptiveCardViewerWebPart;
}(BaseClientSideWebPart));
export default AdaptiveCardViewerWebPart;
//# sourceMappingURL=AdaptiveCardViewerWebPart.js.map