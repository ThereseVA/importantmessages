import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IAdaptiveCardViewerProps {
    context: WebPartContext;
    messageId?: number;
    description?: string;
}
export interface IAdaptiveCardViewerState {
    isLoading: boolean;
    currentView: 'messages' | 'dashboard' | 'creator' | 'diagnostic';
    userRole: 'manager' | 'employee';
    userRoleDetails?: {
        role: 'Employee' | 'Manager' | 'Admin' | 'SuperAdmin';
        method: string;
        isManager: boolean;
    };
    error?: string;
}
/**
 * Adaptive Card Viewer Component
 * Displays adaptive cards based on user role and selected view
 */
export default class AdaptiveCardViewer extends React.Component<IAdaptiveCardViewerProps, IAdaptiveCardViewerState> {
    constructor(props: IAdaptiveCardViewerProps);
    componentDidMount(): Promise<void>;
    render(): React.ReactElement<IAdaptiveCardViewerProps>;
    private renderCurrentView;
}
//# sourceMappingURL=AdaptiveCardViewer.d.ts.map