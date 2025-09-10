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
import styles from './AdaptiveCardComponent.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';
import { HttpClient } from '@microsoft/sp-http';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import * as AdaptiveCards from 'adaptivecards';
import { cardTemplates } from '../models/CardTemplates';
import { TeamsMessageCreator } from './TeamsMessageCreator';
import { ManagerDashboard } from './ManagerDashboard';
var AdaptiveCardComponent = /** @class */ (function (_super) {
    __extends(AdaptiveCardComponent, _super);
    function AdaptiveCardComponent(props) {
        var _this = _super.call(this, props) || this;
        console.log('🏗️ AdaptiveCardComponent constructor called with props:', props);
        _this.cardContainer = React.createRef();
        _this.state = {
            cardData: null,
            loading: false,
            error: null
        };
        console.log('🏗️ AdaptiveCardComponent constructor completed');
        return _this;
    }
    AdaptiveCardComponent.prototype.componentDidMount = function () {
        var _a, _b, _c;
        return __awaiter(this, void 0, void 0, function () {
            var error_1;
            var _this = this;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        console.log('🚀 AdaptiveCardComponent v1.0.34.0 - Component mounted with Teams multi-site support');
                        console.log('📊 Component state:', this.state);
                        console.log('📋 Component props:', this.props);
                        console.log('🔗 Card JSON URL:', this.props.cardJsonUrl);
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 3, , 4]);
                        console.log('🔧 Initializing Enhanced DataService...');
                        return [4 /*yield*/, enhancedDataService.initialize(this.props.context, (_c = (_b = (_a = this.props.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.web) === null || _c === void 0 ? void 0 : _c.absoluteUrl)];
                    case 2:
                        _d.sent();
                        console.log('✅ Enhanced DataService initialized successfully');
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _d.sent();
                        console.error('❌ Error initializing Enhanced DataService:', error_1);
                        return [3 /*break*/, 4];
                    case 4:
                        if (!this.props.cardJsonUrl) {
                            console.log('⚠️ No cardJsonUrl provided, rendering default card');
                            // Render default card if no URL is configured
                            setTimeout(function () { return _this.renderAdaptiveCard(_this.getDefaultCard()); }, 100);
                        }
                        else {
                            console.log('🎯 Loading card from URL:', this.props.cardJsonUrl);
                            // Load card from URL
                            this.loadCard();
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    AdaptiveCardComponent.prototype.componentDidUpdate = function (prevProps) {
        if (prevProps.cardJsonUrl !== this.props.cardJsonUrl && this.props.cardJsonUrl) {
            this.loadCard();
        }
    };
    AdaptiveCardComponent.prototype.loadCard = function () {
        return __awaiter(this, void 0, void 0, function () {
            var cardData, templateName, componentName, assetName, assetError_1, response, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ loading: true, error: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 13, , 14]);
                        cardData = void 0;
                        console.log('🚀🚀🚀 LATEST VERSION 1.0.14.0 - Loading card from:', this.props.cardJsonUrl);
                        console.log('🔧🔧🔧 FETCH API FIX ACTIVE - No more "Failed to fetch" errors - NEW BUNDLE! 🔧🔧🔧');
                        console.log('💥💥💥 v1.0.34.0 CACHE BREAK - TIMESTAMP: ' + Date.now() + ' 💥💥💥');
                        if (!this.props.cardJsonUrl.startsWith('template:')) return [3 /*break*/, 2];
                        templateName = this.props.cardJsonUrl.replace('template:', '');
                        console.log('🎯 FIXED VERSION - Loading template:', templateName);
                        cardData = cardTemplates[templateName];
                        if (!cardData) {
                            throw new Error("Template '".concat(templateName, "' not found"));
                        }
                        return [3 /*break*/, 12];
                    case 2:
                        if (!this.props.cardJsonUrl.startsWith('component:')) return [3 /*break*/, 3];
                        componentName = this.props.cardJsonUrl.replace('component:', '');
                        console.log('🎯 COMPONENT MODE - Loading component:', componentName);
                        this.renderComponent(componentName);
                        return [2 /*return*/];
                    case 3:
                        if (!this.props.cardJsonUrl.startsWith('asset:')) return [3 /*break*/, 8];
                        assetName = this.props.cardJsonUrl.replace('asset:', '');
                        console.log('🎯 FIXED VERSION - Loading asset:', assetName);
                        _a.label = 4;
                    case 4:
                        _a.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, this.loadAssetCard(assetName)];
                    case 5:
                        cardData = _a.sent();
                        console.log('🎯 FIXED VERSION - Asset loaded successfully:', cardData);
                        return [3 /*break*/, 7];
                    case 6:
                        assetError_1 = _a.sent();
                        console.error('🎯 FIXED VERSION - Error loading asset:', assetError_1);
                        throw assetError_1;
                    case 7: return [3 /*break*/, 12];
                    case 8:
                        if (!(this.props.cardJsonUrl.startsWith('http://') || this.props.cardJsonUrl.startsWith('https://'))) return [3 /*break*/, 11];
                        console.log('Loading from URL:', this.props.cardJsonUrl);
                        return [4 /*yield*/, this.props.context.httpClient.get(this.props.cardJsonUrl, HttpClient.configurations.v1)];
                    case 9:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("HTTP ".concat(response.status, ": ").concat(response.statusText));
                        }
                        return [4 /*yield*/, response.json()];
                    case 10:
                        cardData = _a.sent();
                        return [3 /*break*/, 12];
                    case 11: throw new Error("Invalid card source: ".concat(this.props.cardJsonUrl, ". Must be template:, asset:, or a valid HTTP(S) URL."));
                    case 12:
                        this.setState({ cardData: cardData, loading: false });
                        this.renderAdaptiveCard(cardData);
                        return [3 /*break*/, 14];
                    case 13:
                        error_2 = _a.sent();
                        this.setState({
                            error: error_2.message || 'Failed to load Adaptive Card',
                            loading: false
                        });
                        return [3 /*break*/, 14];
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    AdaptiveCardComponent.prototype.loadAssetCard = function (assetName) {
        return __awaiter(this, void 0, void 0, function () {
            var embeddedCards, card;
            return __generator(this, function (_a) {
                embeddedCards = {
                    'sample-card': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.3",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "Sample Adaptive Card",
                                "weight": "Bolder",
                                "size": "Large",
                                "color": "Accent"
                            },
                            {
                                "type": "TextBlock",
                                "text": "This card is loaded from embedded JSON in the SPFx solution.",
                                "wrap": true,
                                "spacing": "Medium"
                            }
                        ]
                    },
                    'project-status': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.3",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "Project Status",
                                "weight": "Bolder",
                                "size": "Large",
                                "color": "Good"
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {
                                        "title": "Status:",
                                        "value": "In Progress"
                                    },
                                    {
                                        "title": "Completion:",
                                        "value": "75%"
                                    }
                                ]
                            }
                        ]
                    },
                    'team-notification': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.3",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "🚀 Team Notification",
                                "weight": "Bolder",
                                "size": "Large",
                                "color": "Accent"
                            },
                            {
                                "type": "TextBlock",
                                "text": "New feature release available! Check out the enhanced Adaptive Cards integration.",
                                "wrap": true,
                                "spacing": "Medium"
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {
                                        "title": "Release Date:",
                                        "value": "July 30, 2025"
                                    },
                                    {
                                        "title": "Version:",
                                        "value": "1.0.9.0"
                                    }
                                ]
                            }
                        ]
                    },
                    'sales-dashboard': {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.3",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "📊 Sales Dashboard",
                                "weight": "Bolder",
                                "size": "Large",
                                "color": "Accent"
                            },
                            {
                                "type": "TextBlock",
                                "text": "Monthly Revenue: $125,000 (↗️ +15%)",
                                "weight": "Bolder",
                                "color": "Good",
                                "spacing": "Medium"
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {
                                        "title": "Active Deals:",
                                        "value": "47"
                                    },
                                    {
                                        "title": "Top Performer:",
                                        "value": "Sarah Johnson - $45,000"
                                    }
                                ]
                            }
                        ]
                    }
                };
                card = embeddedCards[assetName];
                if (!card) {
                    throw new Error("Asset '".concat(assetName, "' not found"));
                }
                // Return the card directly (no fetch needed since it's embedded)
                return [2 /*return*/, Promise.resolve(card)];
            });
        });
    };
    AdaptiveCardComponent.prototype.renderAdaptiveCard = function (cardJson) {
        var _this = this;
        if (!this.cardContainer.current)
            return;
        // Clear previous content
        this.cardContainer.current.innerHTML = '';
        try {
            // Create Adaptive Card instance
            var adaptiveCard = new AdaptiveCards.AdaptiveCard();
            // Set up action handling
            adaptiveCard.onExecuteAction = function (action) {
                _this.handleSubmitAction(action);
            };
            // Parse and render the card
            adaptiveCard.parse(cardJson);
            var renderedCard = adaptiveCard.render();
            if (renderedCard) {
                this.cardContainer.current.appendChild(renderedCard);
            }
        }
        catch (error) {
            console.error('Error rendering Adaptive Card:', error);
            this.setState({ error: 'Failed to render Adaptive Card' });
        }
    };
    AdaptiveCardComponent.prototype.handleSubmitAction = function (action) {
        return __awaiter(this, void 0, void 0, function () {
            var data, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        if (!(action instanceof AdaptiveCards.SubmitAction)) return [3 /*break*/, 6];
                        data = action.data;
                        if (!(data && data.action === 'markAsRead' && data.messageId)) return [3 /*break*/, 4];
                        if (!(typeof data.messageId === 'number')) return [3 /*break*/, 2];
                        return [4 /*yield*/, enhancedDataService.markMessageAsRead(data.messageId)];
                    case 1:
                        _a.sent();
                        console.log("Message ".concat(data.messageId, " marked as read successfully"));
                        // Show success notification
                        this.showSuccessMessage('Message marked as read');
                        return [3 /*break*/, 3];
                    case 2:
                        console.log('Sample card action - would mark message as read:', data);
                        _a.label = 3;
                    case 3: return [3 /*break*/, 5];
                    case 4:
                        console.log('Submit action data:', data);
                        _a.label = 5;
                    case 5: return [3 /*break*/, 7];
                    case 6:
                        console.log('Non-submit action executed:', action);
                        _a.label = 7;
                    case 7: return [3 /*break*/, 9];
                    case 8:
                        error_3 = _a.sent();
                        console.error('Error handling submit action:', error_3);
                        this.showErrorMessage('Failed to process action');
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    AdaptiveCardComponent.prototype.showSuccessMessage = function (message) {
        // Simple success indicator - you could enhance this with a proper notification system
        var successDiv = document.createElement('div');
        successDiv.innerHTML = "\u2713 ".concat(message);
        successDiv.style.cssText = "\n      position: fixed;\n      top: 20px;\n      right: 20px;\n      background: #107c10;\n      color: white;\n      padding: 12px 16px;\n      border-radius: 4px;\n      z-index: 1000;\n      font-family: 'Segoe UI', sans-serif;\n      box-shadow: 0 2px 8px rgba(0,0,0,0.2);\n    ";
        document.body.appendChild(successDiv);
        setTimeout(function () {
            if (successDiv.parentNode) {
                successDiv.parentNode.removeChild(successDiv);
            }
        }, 3000);
    };
    AdaptiveCardComponent.prototype.showErrorMessage = function (message) {
        // Simple error indicator
        var errorDiv = document.createElement('div');
        errorDiv.innerHTML = "\u26A0 ".concat(message);
        errorDiv.style.cssText = "\n      position: fixed;\n      top: 20px;\n      right: 20px;\n      background: #d13438;\n      color: white;\n      padding: 12px 16px;\n      border-radius: 4px;\n      z-index: 1000;\n      font-family: 'Segoe UI', sans-serif;\n      box-shadow: 0 2px 8px rgba(0,0,0,0.2);\n    ";
        document.body.appendChild(errorDiv);
        setTimeout(function () {
            if (errorDiv.parentNode) {
                errorDiv.parentNode.removeChild(errorDiv);
            }
        }, 5000);
    };
    AdaptiveCardComponent.prototype.getDefaultCard = function () {
        return {
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "Welcome to Adaptive Cards!",
                    "size": "Large",
                    "weight": "Bolder"
                },
                {
                    "type": "TextBlock",
                    "text": "This is a sample Adaptive Card showing integration with SharePoint Framework. Configure the web part to load your own card JSON.",
                    "wrap": true
                },
                {
                    "type": "FactSet",
                    "facts": [
                        {
                            "title": "Framework:",
                            "value": "SharePoint Framework"
                        },
                        {
                            "title": "Technology:",
                            "value": "Adaptive Cards"
                        },
                        {
                            "title": "Version:",
                            "value": "1.5"
                        },
                        {
                            "title": "Integration:",
                            "value": "SharePoint Lists & Power Automate"
                        }
                    ]
                }
            ],
            "actions": [
                {
                    "type": "Action.OpenUrl",
                    "title": "Learn More",
                    "url": "https://adaptivecards.io/"
                },
                {
                    "type": "Action.Submit",
                    "title": "Mark as Read",
                    "data": {
                        "action": "markAsRead",
                        "messageId": "sample"
                    }
                }
            ]
        };
    };
    AdaptiveCardComponent.prototype.renderPlaceholder = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.placeholder },
            React.createElement("div", { className: styles.placeholderIcon }, "\uD83D\uDCCB"),
            React.createElement("div", { className: styles.placeholderTitle }, "Configure your Adaptive Card"),
            React.createElement("div", { className: styles.placeholderDescription }, "Please configure the Card JSON URL in the web part properties."),
            React.createElement("button", { className: styles.configureButton, onClick: function () { return _this.props.context.propertyPane.open(); } }, "Configure")));
    };
    AdaptiveCardComponent.prototype.renderTitle = function () {
        var _this = this;
        if (this.props.displayMode === DisplayMode.Edit) {
            return (React.createElement("input", { type: "text", value: this.props.title, onChange: function (e) { return _this.props.updateProperty(e.target.value); }, placeholder: "Enter web part title", style: {
                    fontSize: '18px',
                    fontWeight: 'bold',
                    border: '1px dashed #ccc',
                    padding: '4px 8px',
                    background: 'transparent',
                    width: '100%',
                    marginBottom: '16px'
                } }));
        }
        return this.props.title ? (React.createElement("h2", { style: { marginBottom: '16px', fontSize: '18px', fontWeight: 'bold' } }, escape(this.props.title))) : null;
    };
    AdaptiveCardComponent.prototype.renderComponent = function (componentName) {
        this.setState({ loading: false, error: null });
        // Component rendering will happen in the render method
    };
    AdaptiveCardComponent.prototype.render = function () {
        var _a;
        console.log('🎨 AdaptiveCardComponent.render() called');
        console.log('🎨 Props:', this.props);
        console.log('🎨 State:', this.state);
        console.log('🎨 Card JSON URL:', this.props.cardJsonUrl);
        // Check if we're in component mode
        if ((_a = this.props.cardJsonUrl) === null || _a === void 0 ? void 0 : _a.startsWith('component:')) {
            var componentName = this.props.cardJsonUrl.replace('component:', '');
            console.log('🎨 Component mode detected:', componentName);
            switch (componentName) {
                case 'teams-message-creator':
                    console.log('🎨 Rendering TeamsMessageCreator component');
                    return (React.createElement("div", { className: styles.adaptiveCardComponent },
                        React.createElement(TeamsMessageCreator, { context: this.props.context })));
                case 'manager-dashboard':
                    console.log('🎨 Rendering ManagerDashboard component');
                    // Enhanced data service is globally available
                    return (React.createElement("div", { className: styles.adaptiveCardComponent },
                        React.createElement(ManagerDashboard, null)));
                case 'message-list-diagnostic':
                    console.log('🎨 MessageListDiagnostic component removed');
                    return (React.createElement("div", { className: styles.adaptiveCardComponent },
                        React.createElement("div", null, "MessageListDiagnostic component has been removed as part of cleanup")));
                default:
                    console.log('🎨 Unknown component, rendering error');
                    return (React.createElement("div", { className: styles.adaptiveCardComponent },
                        React.createElement("div", { className: styles.error },
                            React.createElement("div", { className: styles.errorIcon }, "\u26A0\uFE0F"),
                            React.createElement("div", null,
                                "Unknown component: ",
                                componentName))));
            }
        }
        // Only show placeholder in edit mode when no URL is configured AND it's not loading/showing content
        if (this.props.displayMode === DisplayMode.Edit && !this.props.cardJsonUrl && !this.state.cardData && !this.state.loading) {
            console.log('🎨 Rendering placeholder (edit mode, no URL)');
            return (React.createElement("div", { className: styles.adaptiveCardComponent }, this.renderPlaceholder()));
        }
        console.log('🎨 Rendering main component with state-based content');
        return (React.createElement("div", { className: styles.adaptiveCardComponent },
            this.renderTitle(),
            this.state.loading && (React.createElement("div", { className: styles.loading },
                React.createElement("div", { className: styles.spinner }),
                React.createElement("span", null, "Loading Adaptive Card..."))),
            this.state.error && (React.createElement("div", { className: styles.error },
                React.createElement("div", { className: styles.errorIcon }, "\u26A0\uFE0F"),
                React.createElement("div", null,
                    React.createElement("strong", null, "Error loading Adaptive Card:"),
                    React.createElement("br", null),
                    this.state.error,
                    React.createElement("br", null),
                    React.createElement("small", null,
                        "URL: ",
                        this.props.cardJsonUrl)))),
            !this.state.loading && !this.state.error && (React.createElement("div", { ref: this.cardContainer, className: styles.cardContainer }))));
    };
    return AdaptiveCardComponent;
}(React.Component));
export { AdaptiveCardComponent };
//# sourceMappingURL=AdaptiveCardComponent.js.map