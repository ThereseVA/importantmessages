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
import * as React from 'react';
import { useState, useEffect } from 'react';
import { PrimaryButton, DefaultButton, TextField, Dropdown, MessageBar, MessageBarType, Label, Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { EnhancedTeamsService } from '../../../services/EnhancedTeamsService';
import { enhancedDataService } from '../../../services/EnhancedDataService';
import { SiteSelector } from './SiteSelector';
export var TeamsMessageCreator = function (props) {
    var _a, _b, _c, _d, _e, _f;
    console.log('🎯 TeamsMessageCreator component started');
    console.log('🎯 Props received:', props);
    console.log('🎯 Context available:', !!props.context);
    // Manager permission state
    var _g = useState(null), isManager = _g[0], setIsManager = _g[1];
    var _h = useState(true), isCheckingPermissions = _h[0], setIsCheckingPermissions = _h[1];
    // Initialize with current SharePoint context if available
    var _j = useState(((_c = (_b = (_a = props.context) === null || _a === void 0 ? void 0 : _a.pageContext) === null || _b === void 0 ? void 0 : _b.web) === null || _c === void 0 ? void 0 : _c.absoluteUrl) || 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken'), currentSite = _j[0], setCurrentSite = _j[1];
    var _k = useState(((_f = (_e = (_d = props.context) === null || _d === void 0 ? void 0 : _d.pageContext) === null || _e === void 0 ? void 0 : _e.web) === null || _f === void 0 ? void 0 : _f.title) || 'Current Site'), currentSiteName = _k[0], setCurrentSiteName = _k[1];
    var _l = useState({
        title: '',
        content: '',
        priority: 'Medium',
        targetAudience: 'Teams Channel',
        expiryDays: '7',
        distributionChannels: [],
        useEmailIntegration: false // New option for email-based Teams integration
    }), formData = _l[0], setFormData = _l[1];
    var _m = useState(false), isSubmitting = _m[0], setIsSubmitting = _m[1];
    var _o = useState(null), result = _o[0], setResult = _o[1];
    var _p = useState(''), webhookUrls = _p[0], setWebhookUrls = _p[1];
    console.log('🎯 TeamsMessageCreator state initialized');
    console.log('🎯 Form data:', formData);
    // Check manager permissions on component mount
    useEffect(function () {
        var checkManagerPermissions = function () { return __awaiter(void 0, void 0, void 0, function () {
            var initError_1, managerStatus, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!props.context) {
                            console.warn('🎯 No context available for permission check');
                            setIsCheckingPermissions(false);
                            setIsManager(false);
                            return [2 /*return*/];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 7, , 8]);
                        console.log('🎯 Checking manager permissions...');
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 4, , 5]);
                        return [4 /*yield*/, enhancedDataService.initialize(props.context)];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        initError_1 = _a.sent();
                        console.warn('🎯 Service already initialized or initialization failed:', initError_1);
                        return [3 /*break*/, 5];
                    case 5: return [4 /*yield*/, enhancedDataService.isCurrentUserManager()];
                    case 6:
                        managerStatus = _a.sent();
                        console.log('🎯 Manager status:', managerStatus);
                        setIsManager(managerStatus);
                        setIsCheckingPermissions(false);
                        return [3 /*break*/, 8];
                    case 7:
                        error_1 = _a.sent();
                        console.error('🎯 Error checking manager permissions:', error_1);
                        setIsManager(false);
                        setIsCheckingPermissions(false);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        }); };
        checkManagerPermissions();
    }, [props.context]);
    // Rich text editor functions - safer implementation
    var contentRef = React.useRef(null);
    var formatText = function (command, value) {
        try {
            // Only execute if the content ref is available and focused
            if (contentRef.current && document.activeElement === contentRef.current) {
                document.execCommand(command, false, value);
            }
        }
        catch (error) {
            console.warn('Format command failed:', error);
        }
    };
    var handleContentChange = function (event) {
        var content = event.currentTarget.innerHTML;
        setFormData(__assign(__assign({}, formData), { content: content }));
    };
    var priorityOptions = [
        { key: 'High', text: '🚨 High Priority' },
        { key: 'Medium', text: '⚠️ Medium Priority' },
        { key: 'Low', text: 'ℹ️ Low Priority' }
    ];
    var audienceOptions = [
        { key: 'All Teams', text: '� All Teams Channels' },
        { key: 'General Channel', text: '🏢 General Channel' },
        { key: 'Leadership Team', text: '👔 Leadership Team Chat' },
        { key: 'IT Support Channel', text: '💻 IT Support Channel' },
        { key: 'Medical Staff', text: '🏥 Medical Staff Channel' },
        { key: 'Nursing Team', text: '�‍⚕️ Nursing Team Chat' },
        { key: 'Administration', text: '📋 Administration Channel' },
        { key: 'Emergency Response', text: '🚨 Emergency Response Channel' },
        { key: 'Department Heads', text: '🎯 Department Heads Chat' },
        { key: 'Custom Teams', text: '✏️ Custom Teams/Channels' }
    ];
    var expiryOptions = [
        { key: '1', text: '1 Day' },
        { key: '3', text: '3 Days' },
        { key: '7', text: '1 Week' },
        { key: '14', text: '2 Weeks' },
        { key: '30', text: '1 Month' }
    ];
    var handleSubmit = function () { return __awaiter(void 0, void 0, void 0, function () {
        var contextSiteUrl, expiryDate, isTeamsContext, newMessage, messageId, fullMessage, emailResult, htmlNotification, shareLink, copyPasteMessage, error_2, errorMessage;
        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m, _o, _p, _q, _r;
        return __generator(this, function (_s) {
            switch (_s.label) {
                case 0:
                    if (!currentSite) {
                        setResult({ type: 'error', message: '❌ Please select a SharePoint site first' });
                        return [2 /*return*/];
                    }
                    if (!formData.title.trim() || !formData.content.trim()) {
                        setResult({ type: 'error', message: '❌ Please fill in title and content' });
                        return [2 /*return*/];
                    }
                    setIsSubmitting(true);
                    setResult({ type: 'info', message: "\uD83D\uDCE4 Creating message in ".concat(currentSiteName, "...") });
                    _s.label = 1;
                case 1:
                    _s.trys.push([1, 13, 14, 15]);
                    if (!!enhancedDataService.getCurrentUser()) return [3 /*break*/, 3];
                    return [4 /*yield*/, enhancedDataService.initialize(props.context, currentSite)];
                case 2:
                    _s.sent();
                    return [3 /*break*/, 4];
                case 3:
                    // Update site URL if changed
                    enhancedDataService.setSharePointSiteUrl(currentSite);
                    _s.label = 4;
                case 4:
                    console.log('🔍 DEBUG: Enhanced Data Service initialized');
                    console.log('🔍 DEBUG: Current site URL:', currentSite);
                    console.log('🔍 DEBUG: Current site name:', currentSiteName);
                    if (props.context) {
                        console.log('🔍 DEBUG: SPFx web URL:', (_b = (_a = props.context.pageContext) === null || _a === void 0 ? void 0 : _a.web) === null || _b === void 0 ? void 0 : _b.absoluteUrl);
                        console.log('🔍 DEBUG: SPFx web title:', (_d = (_c = props.context.pageContext) === null || _c === void 0 ? void 0 : _c.web) === null || _d === void 0 ? void 0 : _d.title);
                        console.log('🔍 DEBUG: SPFx user:', (_f = (_e = props.context.pageContext) === null || _e === void 0 ? void 0 : _e.user) === null || _f === void 0 ? void 0 : _f.displayName);
                        contextSiteUrl = (_h = (_g = props.context.pageContext) === null || _g === void 0 ? void 0 : _g.web) === null || _h === void 0 ? void 0 : _h.absoluteUrl;
                        console.log('🔍 CRITICAL DEBUG:');
                        console.log('   - currentSite state:', currentSite);
                        console.log('   - SPFx context site:', contextSiteUrl);
                        console.log('   - Are they the same?', currentSite === contextSiteUrl);
                        // IMPORTANT: Always use the enhanced data service
                        console.log('🔧 Enhanced Data Service configured for site:', currentSite);
                        // Check if we're trying to access a different site than the current context
                        if (currentSite && contextSiteUrl && !currentSite.startsWith(contextSiteUrl) && !contextSiteUrl.startsWith(currentSite)) {
                            console.warn('⚠️ CROSS-SITE ACCESS DETECTED:');
                            console.warn('   Context site:', contextSiteUrl);
                            console.warn('   Target site:', currentSite);
                            console.warn('   This may cause 403 Forbidden errors!');
                            setResult({
                                type: 'error',
                                message: "\u274C Cross-site access detected!\n\nContext site: ".concat(contextSiteUrl, "\nTarget site: ").concat(currentSite, "\n\n\uD83D\uDCA1 You may not have permission to access the target site from this context. Try:\n1. Opening the web part directly on the target site\n2. Using the same site for both context and target\n3. Ensuring you have proper cross-site permissions")
                            });
                            return [2 /*return*/];
                        }
                    }
                    else {
                        console.warn('⚠️ No SPFx context available - this may cause authentication issues');
                        setResult({
                            type: 'error',
                            message: '❌ No SharePoint context available!\n\nThis component requires SPFx context to access SharePoint. Make sure:\n1. The web part is added to a SharePoint page\n2. You\'re not viewing in preview mode\n3. The page has fully loaded'
                        });
                        return [2 /*return*/];
                    }
                    expiryDate = new Date();
                    expiryDate.setDate(expiryDate.getDate() + parseInt(formData.expiryDays));
                    isTeamsContext = !props.context || window.location.href.includes('teams.microsoft.com');
                    newMessage = {
                        Title: formData.title,
                        MessageContent: formData.content,
                        Priority: formData.priority,
                        TargetAudience: formData.targetAudience,
                        ExpiryDate: expiryDate,
                        Source: isTeamsContext ? 'Teams' : 'SharePoint'
                    };
                    console.log('📝 Creating message with data:', newMessage);
                    console.log('🎯 Target site:', currentSite);
                    console.log('🔗 SharePoint context site:', (_k = (_j = props.context.pageContext) === null || _j === void 0 ? void 0 : _j.web) === null || _k === void 0 ? void 0 : _k.absoluteUrl);
                    console.log('👤 Current user:', (_m = (_l = props.context.pageContext) === null || _l === void 0 ? void 0 : _l.user) === null || _m === void 0 ? void 0 : _m.displayName);
                    console.log('📧 User email:', (_p = (_o = props.context.pageContext) === null || _o === void 0 ? void 0 : _o.user) === null || _p === void 0 ? void 0 : _p.email);
                    console.log('🌐 Window location:', window.location.href);
                    // Validate that we have proper SharePoint context
                    if (!((_r = (_q = props.context.pageContext) === null || _q === void 0 ? void 0 : _q.web) === null || _r === void 0 ? void 0 : _r.absoluteUrl)) {
                        setResult({
                            type: 'error',
                            message: '❌ Invalid SharePoint context!\n\nThe web context is not available. This usually means:\n1. The component is not running in a proper SharePoint context\n2. The page hasn\'t fully loaded\n3. There\'s a permissions issue with the current site'
                        });
                        return [2 /*return*/];
                    }
                    return [4 /*yield*/, enhancedDataService.createMessage(newMessage)];
                case 5:
                    messageId = _s.sent();
                    console.log('✅ Message created with ID:', messageId);
                    if (!messageId || messageId <= 0) {
                        setResult({
                            type: 'error',
                            message: '❌ Message creation failed!\n\nThe message was not created successfully. Check:\n1. SharePoint list "Important Messages" exists\n2. You have contribute permissions\n3. Required fields are properly configured\n4. Browser console for detailed error messages'
                        });
                        return [2 /*return*/];
                    }
                    return [4 /*yield*/, enhancedDataService.getMessageById(messageId)];
                case 6:
                    fullMessage = _s.sent();
                    console.log('📄 Retrieved full message:', fullMessage);
                    if (!formData.useEmailIntegration) return [3 /*break*/, 8];
                    // 📧 NEW: Enhanced Teams integration using Graph API
                    console.log('📧 Using enhanced Teams integration...');
                    return [4 /*yield*/, EnhancedTeamsService.distributeToAccessibleChannels(fullMessage)];
                case 7:
                    emailResult = _s.sent();
                    if (emailResult.success === 0) {
                        setResult({
                            type: 'error',
                            message: "\u274C No Teams channels accessible!\n\nMessage created in SharePoint (ID: ".concat(messageId, ") but no accessible Teams channels found.\n\n\uD83D\uDCA1 Make sure you have access to Teams channels or check permissions.")
                        });
                    }
                    else if (emailResult.failed === 0) {
                        setResult({
                            type: 'success',
                            message: "\u2705 Message created and sent to Teams!\n\uD83D\uDCCA Sent to ".concat(emailResult.success, " Teams channels\n\uD83D\uDCCB Message ID: ").concat(messageId)
                        });
                    }
                    else {
                        setResult({
                            type: 'success',
                            message: "\u26A0\uFE0F Partial success!\n\uD83D\uDCCA Sent to ".concat(emailResult.success, " channels, ").concat(emailResult.failed, " failed\n\uD83D\uDCCB Message ID: ").concat(messageId, "\n\n\uD83D\uDCA1 Check Teams permissions and channel access.")
                        });
                    }
                    return [3 /*break*/, 12];
                case 8:
                    if (!webhookUrls.trim()) return [3 /*break*/, 11];
                    // 🔗 Enhanced Teams distribution using Graph integration
                    console.log('🔗 Using enhanced Teams integration...');
                    return [4 /*yield*/, EnhancedTeamsService.createNotification(fullMessage)];
                case 9:
                    htmlNotification = _s.sent();
                    return [4 /*yield*/, EnhancedTeamsService.createShareableLink(fullMessage)];
                case 10:
                    shareLink = _s.sent();
                    copyPasteMessage = EnhancedTeamsService.createCopyPasteMessage(fullMessage);
                    console.log('✅ Created alternative distribution content');
                    setResult({
                        type: 'success',
                        message: "\u2705 Message created with enhanced distribution options!\n\n\uFFFD Message ID: ".concat(messageId, "\n\uD83D\uDD17 Shareable link created\n\uD83D\uDCDD Copy-paste message ready\n\uD83D\uDCA1 Use the dashboard to view and share the message")
                    });
                    return [3 /*break*/, 12];
                case 11:
                    setResult({
                        type: 'success',
                        message: "\u2705 Message created in SharePoint!\nMessage ID: ".concat(messageId, "\n\uD83D\uDCA1 Enable email integration or add webhook URLs to distribute to Teams")
                    });
                    _s.label = 12;
                case 12:
                    // Reset form
                    setFormData({
                        title: '',
                        content: '',
                        priority: 'Medium',
                        targetAudience: 'Teams Channel',
                        expiryDays: '7',
                        distributionChannels: [],
                        useEmailIntegration: false
                    });
                    // Clear the rich text editor safely
                    if (contentRef.current) {
                        contentRef.current.innerHTML = '';
                    }
                    setWebhookUrls('');
                    if (props.onMessageCreated) {
                        props.onMessageCreated(messageId);
                    }
                    return [3 /*break*/, 15];
                case 13:
                    error_2 = _s.sent();
                    console.error('❌ Error creating message:', error_2);
                    console.error('📋 Form data was:', formData);
                    console.error('🎯 Target site was:', currentSite);
                    console.error('💾 Message data was:', {
                        Title: formData.title,
                        MessageContent: formData.content,
                        Priority: formData.priority,
                        TargetAudience: formData.targetAudience
                    });
                    errorMessage = "\u274C Failed to create message: ".concat(error_2.message);
                    if (error_2.message.includes('404') || error_2.message.includes('Not Found')) {
                        errorMessage += '\n\n💡 Possible issues:\n• SharePoint list "Important Messages" may not exist\n• You may not have access to the selected site\n• The list may have a different name';
                    }
                    else if (error_2.message.includes('400') || error_2.message.includes('Bad Request')) {
                        errorMessage += '\n\n💡 Possible issues:\n• Required field may be missing from SharePoint list\n• Field types may not match\n• Data validation failed';
                    }
                    else if (error_2.message.includes('403') || error_2.message.includes('Forbidden')) {
                        errorMessage += '\n\n💡 Possible issues:\n• You don\'t have permission to add items to the list\n• The site may require additional permissions';
                    }
                    setResult({ type: 'error', message: errorMessage });
                    return [3 /*break*/, 15];
                case 14:
                    setIsSubmitting(false);
                    return [7 /*endfinally*/];
                case 15: return [2 /*return*/];
            }
        });
    }); };
    // Authentication test functions
    var runAuthTest = function () { return __awaiter(void 0, void 0, void 0, function () {
        var user, error_3;
        var _a, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    if (!props.context) {
                        setResult({ type: 'error', message: '❌ No SPFx context available for authentication test' });
                        return [2 /*return*/];
                    }
                    setResult({ type: 'info', message: '🔍 Running SharePoint authentication test...' });
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, enhancedDataService.initialize(props.context, currentSite)];
                case 2:
                    _c.sent();
                    user = enhancedDataService.getCurrentUser();
                    if (user) {
                        setResult({ type: 'success', message: "\u2705 Authentication test passed!\nUser: ".concat(((_a = user.spfx) === null || _a === void 0 ? void 0 : _a.displayName) || ((_b = user.spfx) === null || _b === void 0 ? void 0 : _b.email) || 'Unknown') });
                    }
                    else {
                        setResult({ type: 'error', message: '❌ Authentication test failed - could not get current user' });
                    }
                    return [3 /*break*/, 4];
                case 3:
                    error_3 = _c.sent();
                    setResult({ type: 'error', message: "\u274C Authentication test failed!\n".concat(error_3.message) });
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var testMessageCreation = function () { return __awaiter(void 0, void 0, void 0, function () {
        var testMessage, messageId, error_4;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!props.context) {
                        setResult({ type: 'error', message: '❌ No SPFx context available for message creation test' });
                        return [2 /*return*/];
                    }
                    setResult({ type: 'info', message: '📝 Testing message creation directly...' });
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 4, , 5]);
                    return [4 /*yield*/, enhancedDataService.initialize(props.context, currentSite)];
                case 2:
                    _a.sent();
                    testMessage = {
                        Title: 'Test Message',
                        MessageContent: 'This is a test message to verify functionality.',
                        Priority: 'Medium',
                        TargetAudience: 'Test',
                        ExpiryDate: new Date(Date.now() + 24 * 60 * 60 * 1000)
                    };
                    return [4 /*yield*/, enhancedDataService.createMessage(testMessage)];
                case 3:
                    messageId = _a.sent();
                    setResult({ type: 'success', message: "\u2705 Message creation test passed!\nMessage ID: ".concat(messageId) });
                    return [3 /*break*/, 5];
                case 4:
                    error_4 = _a.sent();
                    setResult({ type: 'error', message: "\u274C Message creation test failed!\n".concat(error_4.message) });
                    return [3 /*break*/, 5];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var handleQuickTemplate = function (template) {
        var templates = {
            urgent: {
                title: '🚨 Verksamhetskritisk Information',
                content: '<p><strong>Detta är verksamhetskritisk information</strong> som kräver <em>omedelbar uppmärksamhet</em>.</p><p style="color: #d73a49;">Vänligen granska och vidta nödvändiga åtgärder.</p>',
                priority: 'High',
                targetAudience: 'Teams Channel',
                expiryDays: '1',
                distributionChannels: []
            },
            maintenance: {
                title: '🔧 Viktig information!',
                content: '<p><strong>Viktig information</strong> som berör verksamheten.</p><ul><li>Läs igenom denna information noggrant</li><li>Kontakta ansvarig vid frågor</li></ul>',
                priority: 'Medium',
                targetAudience: 'Chat Group',
                expiryDays: '3',
                distributionChannels: []
            },
            announcement: {
                title: '📢 Notera',
                content: '<p>Information som är bra att känna till.</p><p style="color: #0366d6;"><em>Läs igenom när du har tid.</em></p>',
                priority: 'Low',
                targetAudience: 'Teams Channel',
                expiryDays: '7',
                distributionChannels: []
            },
            routine: {
                title: '📢 Uppdaterad/Ny Rutin',
                content: '<p style="color: #0366d6;"><strong>Ny eller uppdaterad rutin</strong> har implementerats.</p><p style="color: #0366d6;"><em>Vänligen läs igenom och följ de nya riktlinjerna.</em></p><ul><li style="color: #0366d6;">Granska rutinändringarna</li><li style="color: #0366d6;">Implementera i dagligt arbete</li><li style="color: #0366d6;">Kontakta ansvarig vid frågor</li></ul>',
                priority: 'Low',
                targetAudience: 'Teams Channel',
                expiryDays: '7',
                distributionChannels: []
            }
        };
        var template_data = templates[template];
        // Update all form fields with template data
        setFormData({
            title: template_data.title,
            content: template_data.content,
            priority: template_data.priority,
            targetAudience: template_data.targetAudience,
            expiryDays: template_data.expiryDays,
            distributionChannels: template_data.distributionChannels,
            useEmailIntegration: false
        });
        // Update the rich text editor content safely
        if (contentRef.current) {
            contentRef.current.innerHTML = template_data.content;
        }
    };
    var handleSiteSelected = function (siteUrl, siteName) {
        setCurrentSite(siteUrl);
        setCurrentSiteName(siteName);
        setResult({ type: 'info', message: "\u2705 Connected to ".concat(siteName) });
    };
    return (React.createElement("div", { style: { padding: '20px', maxWidth: '800px' } },
        React.createElement("h2", null, "\uD83D\uDCDD Create Message from Teams"),
        React.createElement("p", null, "Create and distribute important messages directly from Microsoft Teams"),
        console.log('🎯 TeamsMessageCreator render() called - Component is rendering!'),
        isCheckingPermissions && (React.createElement("div", { style: { padding: '20px', textAlign: 'center' } },
            React.createElement(Spinner, { size: SpinnerSize.large, label: "Checking permissions..." }),
            React.createElement("p", { style: { marginTop: '10px', color: '#666' } }, "Verifying your manager access from SharePoint list..."))),
        !isCheckingPermissions && isManager === false && (React.createElement("div", { style: { padding: '20px', textAlign: 'center' } },
            React.createElement(MessageBar, { messageBarType: MessageBarType.blocked, isMultiline: true },
                React.createElement("h3", null, "\uD83D\uDD12 Access Restricted"),
                React.createElement("p", null,
                    React.createElement("strong", null, "Message creation is restricted to managers only.")),
                React.createElement("p", null, "You are not currently listed as a manager in the SharePoint Managers list. If you believe this is an error, please contact your administrator."),
                React.createElement("div", { style: { marginTop: '15px', padding: '10px', backgroundColor: '#fff3cd', borderRadius: '4px' } },
                    React.createElement("strong", null, "How manager access is determined:"),
                    React.createElement("ul", { style: { textAlign: 'left', marginTop: '8px' } },
                        React.createElement("li", null, "Your email must be listed in the \"Managers\" SharePoint list"),
                        React.createElement("li", null, "Your entry must have \"Is Active\" set to \"Yes\""),
                        React.createElement("li", null, "Contact HR or IT to be added to the managers list")))))),
        !isCheckingPermissions && isManager === true && (React.createElement(React.Fragment, null,
            React.createElement(MessageBar, { messageBarType: MessageBarType.success, isMultiline: false, dismissButtonAriaLabel: "Close" }, "\u2705 Manager access confirmed. You can create and distribute messages."),
            React.createElement(SiteSelector, { onSiteSelected: handleSiteSelected, currentSite: currentSite }),
            currentSite && (React.createElement(React.Fragment, null,
                React.createElement("div", { style: { marginBottom: '20px', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '8px' } },
                    React.createElement("h4", null, "\u26A1 Quick Templates:"),
                    React.createElement("div", { style: { display: 'flex', gap: '10px', flexWrap: 'wrap' } },
                        React.createElement(DefaultButton, { text: "\uD83D\uDEA8 Verksamhetskritisk", onClick: function () { return handleQuickTemplate('urgent'); } }),
                        React.createElement(DefaultButton, { text: "\uD83D\uDD27 Viktig information", onClick: function () { return handleQuickTemplate('maintenance'); } }),
                        React.createElement(DefaultButton, { text: "\uD83D\uDCE2 Notera", onClick: function () { return handleQuickTemplate('announcement'); } }),
                        React.createElement(DefaultButton, { text: "\uD83D\uDCE2 Uppdaterad/Ny Rutin", onClick: function () { return handleQuickTemplate('routine'); } }))),
                React.createElement("div", { style: { display: 'grid', gap: '15px' } },
                    React.createElement(TextField, { label: "\uD83D\uDCCB Message Title", value: formData.title, onChange: function (_, value) { return setFormData(__assign(__assign({}, formData), { title: value || '' })); }, placeholder: "Enter a clear, descriptive title", required: true }),
                    React.createElement("div", { style: { marginBottom: '15px' } },
                        React.createElement(Label, { required: true }, "\uD83D\uDCDD Message Content"),
                        React.createElement("div", { style: {
                                border: '1px solid #d0d7de',
                                borderBottom: 'none',
                                padding: '8px',
                                backgroundColor: '#f6f8fa',
                                display: 'flex',
                                gap: '4px',
                                flexWrap: 'wrap'
                            } },
                            React.createElement("button", { type: "button", onClick: function () { return formatText('bold'); }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } },
                                React.createElement("strong", null, "B")),
                            React.createElement("button", { type: "button", onClick: function () { return formatText('italic'); }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } },
                                React.createElement("em", null, "I")),
                            React.createElement("button", { type: "button", onClick: function () { return formatText('underline'); }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } },
                                React.createElement("u", null, "U")),
                            React.createElement("select", { onChange: function (e) { return formatText('fontSize', e.target.value); }, style: { padding: '4px', border: '1px solid #ccc' } },
                                React.createElement("option", { value: "" }, "Font Size"),
                                React.createElement("option", { value: "1" }, "Small"),
                                React.createElement("option", { value: "3" }, "Normal"),
                                React.createElement("option", { value: "5" }, "Large"),
                                React.createElement("option", { value: "7" }, "Extra Large")),
                            React.createElement("input", { type: "color", onChange: function (e) { return formatText('foreColor', e.target.value); }, style: { width: '30px', height: '26px', border: '1px solid #ccc' }, title: "Text Color" }),
                            React.createElement("input", { type: "color", onChange: function (e) { return formatText('backColor', e.target.value); }, style: { width: '30px', height: '26px', border: '1px solid #ccc' }, title: "Background Color" }),
                            React.createElement("button", { type: "button", onClick: function () { return formatText('insertUnorderedList'); }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } }, "\u2022 List"),
                            React.createElement("button", { type: "button", onClick: function () { return formatText('insertOrderedList'); }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } }, "1. List"),
                            React.createElement("button", { type: "button", onClick: function () {
                                    var url = prompt('Enter URL:');
                                    if (url)
                                        formatText('createLink', url);
                                }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } }, "\uD83D\uDD17 Link"),
                            React.createElement("button", { type: "button", onClick: function () {
                                    var tableHtml = '<table border="1" style="border-collapse: collapse; width: 100%;"><tr><td style="padding: 8px;">Cell 1</td><td style="padding: 8px;">Cell 2</td></tr><tr><td style="padding: 8px;">Cell 3</td><td style="padding: 8px;">Cell 4</td></tr></table>';
                                    formatText('insertHTML', tableHtml);
                                }, style: { padding: '4px 8px', border: '1px solid #ccc', background: '#fff' } }, "\uD83D\uDCCA Table")),
                        React.createElement("div", { ref: contentRef, contentEditable: true, onInput: handleContentChange, dangerouslySetInnerHTML: { __html: formData.content }, dir: "ltr", lang: "sv-SE", style: {
                                border: '1px solid #d0d7de',
                                minHeight: '120px',
                                padding: '12px',
                                backgroundColor: '#fff',
                                outline: 'none',
                                fontSize: '14px',
                                lineHeight: '1.5',
                                fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
                                direction: 'ltr',
                                textAlign: 'left',
                                unicodeBidi: 'embed'
                            }, placeholder: "Enter your message content with rich formatting..." }),
                        React.createElement("div", { style: { fontSize: '12px', color: '#666', marginTop: '4px' } }, "\uD83D\uDCA1 Use the toolbar above to format text, add links, tables, and more")),
                    React.createElement("div", { style: { display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '15px' } },
                        React.createElement(Dropdown, { label: "\u26A1 Priority", selectedKey: formData.priority, onChange: function (_, option) { return setFormData(__assign(__assign({}, formData), { priority: (option === null || option === void 0 ? void 0 : option.key) || 'Medium' })); }, options: priorityOptions }),
                        React.createElement(Dropdown, { label: "\uD83D\uDC65 Target Audience", selectedKey: formData.targetAudience, onChange: function (_, option) { return setFormData(__assign(__assign({}, formData), { targetAudience: (option === null || option === void 0 ? void 0 : option.key) || 'Teams Channel' })); }, options: audienceOptions }),
                        React.createElement(Dropdown, { label: "\uD83D\uDCC5 Expires In", selectedKey: formData.expiryDays, onChange: function (_, option) { return setFormData(__assign(__assign({}, formData), { expiryDays: (option === null || option === void 0 ? void 0 : option.key) || '7' })); }, options: expiryOptions })),
                    React.createElement("div", { style: { marginTop: '20px', padding: '15px', backgroundColor: '#f8f9fa', borderRadius: '8px', border: '1px solid #e1e5e9' } },
                        React.createElement("h4", { style: { margin: '0 0 10px 0', color: '#323130' } }, "\uD83D\uDCE7 Teams Integration Method"),
                        React.createElement("div", { style: { display: 'flex', gap: '20px', marginBottom: '15px' } },
                            React.createElement("label", { style: { display: 'flex', alignItems: 'center', cursor: 'pointer' } },
                                React.createElement("input", { type: "radio", name: "teamsIntegration", checked: formData.useEmailIntegration, onChange: function () { return setFormData(__assign(__assign({}, formData), { useEmailIntegration: true })); }, style: { marginRight: '8px' } }),
                                React.createElement("span", null,
                                    "\uD83D\uDCE7 ",
                                    React.createElement("strong", null, "Email Integration"),
                                    " (Easy - uses SharePoint list)")),
                            React.createElement("label", { style: { display: 'flex', alignItems: 'center', cursor: 'pointer' } },
                                React.createElement("input", { type: "radio", name: "teamsIntegration", checked: !formData.useEmailIntegration, onChange: function () { return setFormData(__assign(__assign({}, formData), { useEmailIntegration: false })); }, style: { marginRight: '8px' } }),
                                React.createElement("span", null,
                                    "\uD83D\uDD17 ",
                                    React.createElement("strong", null, "Webhook Integration"),
                                    " (Manual setup required)"))),
                        formData.useEmailIntegration ? (React.createElement("div", { style: { backgroundColor: '#fff', padding: '10px', borderRadius: '4px', border: '1px solid #d1d9e0' } },
                            React.createElement("p", { style: { margin: '0', fontSize: '14px', color: '#605e5c' } },
                                "\u2705 ",
                                React.createElement("strong", null, "Automatic sending to configured Teams channels"),
                                React.createElement("br", null),
                                "\uD83D\uDCCB Channels are configured in the ",
                                React.createElement("a", { href: "https://gustafkliniken.sharepoint.com/sites/Gustafkliniken/Lists/TeamsChannels/AllItems.aspx", target: "_blank", rel: "noopener noreferrer" }, "TeamsChannels SharePoint list"),
                                React.createElement("br", null),
                                "\uD83C\uDFAF Messages will be sent based on priority and department filters"))) : (React.createElement("div", { style: { backgroundColor: '#fff', padding: '10px', borderRadius: '4px', border: '1px solid #d1d9e0' } },
                            React.createElement("p", { style: { margin: '0 0 10px 0', fontSize: '14px', color: '#605e5c' } },
                                "\uD83D\uDD17 ",
                                React.createElement("strong", null, "Manual webhook setup required"),
                                React.createElement("br", null),
                                "\uD83D\uDCA1 Get webhook URLs from Teams channels (Channel \u2192 ... \u2192 Connectors \u2192 Incoming Webhook)")))),
                    !formData.useEmailIntegration && (React.createElement(TextField, { label: "\uD83D\uDD17 Teams Webhook URLs (one per line)", value: webhookUrls, onChange: function (_, value) { return setWebhookUrls(value || ''); }, placeholder: "https://outlook.office.com/webhook/channel1...\nhttps://outlook.office.com/webhook/channel2...", multiline: true, rows: 3, description: "Paste webhook URLs from Teams channels where you want to distribute this message" }))),
                React.createElement("div", { style: { marginTop: '20px', display: 'flex', gap: '10px' } },
                    React.createElement(PrimaryButton, { text: "\uD83D\uDCE4 Create & Distribute", onClick: handleSubmit, disabled: isSubmitting || !formData.title.trim() || !formData.content.trim() }),
                    React.createElement(DefaultButton, { text: "\uD83D\uDCBE Save to SharePoint Only", onClick: handleSubmit, disabled: isSubmitting })),
                result && (React.createElement(MessageBar, { messageBarType: result.type === 'success' ? MessageBarType.success :
                        result.type === 'error' ? MessageBarType.error : MessageBarType.info, styles: { root: { marginTop: '20px' } } },
                    React.createElement("pre", { style: { whiteSpace: 'pre-wrap', fontFamily: 'inherit' } }, result.message))),
                props.context && (React.createElement("div", { style: { marginTop: '20px', padding: '15px', backgroundColor: '#fff8e6', borderRadius: '8px', border: '1px solid #ffd700' } },
                    React.createElement("h4", null, "\uD83D\uDD27 Debugging Tools"),
                    React.createElement("p", null, "If message creation is failing, use these tests to diagnose the issue:"),
                    React.createElement("div", { style: { display: 'flex', gap: '10px', flexWrap: 'wrap' } },
                        React.createElement(DefaultButton, { text: "\uD83D\uDD0D Test Authentication", onClick: runAuthTest, disabled: isSubmitting }),
                        React.createElement(DefaultButton, { text: "\uD83D\uDCDD Test Message Creation", onClick: testMessageCreation, disabled: isSubmitting })),
                    React.createElement("div", { style: { fontSize: '12px', color: '#666', marginTop: '8px' } }, "\uD83D\uDCA1 These tests will check if you have proper access to SharePoint and can create messages"))),
                React.createElement("div", { style: { marginTop: '30px', padding: '15px', backgroundColor: '#e8f4fd', borderRadius: '8px' } },
                    React.createElement("h4", null, "\uD83D\uDCA1 How to Use:"),
                    React.createElement("ol", null,
                        React.createElement("li", null,
                            React.createElement("strong", null, "\uD83C\uDFAF From Teams Channel:"),
                            " Get webhook URL (Channel \u2192 \u22EF \u2192 Connectors \u2192 Incoming Webhook)"),
                        React.createElement("li", null,
                            React.createElement("strong", null, "\uD83D\uDCDD Create Message:"),
                            " Fill out the form above with your message details"),
                        React.createElement("li", null,
                            React.createElement("strong", null, "\uD83D\uDCE4 Distribute:"),
                            " Message goes to SharePoint + selected Teams channels"),
                        React.createElement("li", null,
                            React.createElement("strong", null, "\uD83D\uDCCA Track:"),
                            " View read confirmations in the dashboard")),
                    React.createElement("h4", null, "\uD83D\uDD04 Integration Options:"),
                    React.createElement("ul", null,
                        React.createElement("li", null,
                            React.createElement("strong", null, "Teams Tab:"),
                            " Add this as a tab in your Teams channel"),
                        React.createElement("li", null,
                            React.createElement("strong", null, "Teams Bot:"),
                            " Create a bot for conversational message creation"),
                        React.createElement("li", null,
                            React.createElement("strong", null, "Power Automate:"),
                            " Trigger from Teams messages or reactions"),
                        React.createElement("li", null,
                            React.createElement("strong", null, "Teams App:"),
                            " Package as a full Teams application")))))))));
};
//# sourceMappingURL=TeamsMessageCreator.js.map