import * as React from 'react';
import { IDashboardProps } from './IDashboardProps';
import { IMessage } from '../../../services/EnhancedDataService';
export interface IDashboardState {
    messages: IMessage[];
    filteredMessages: IMessage[];
    loading: boolean;
    error: string | null;
    lastRefresh: Date | null;
    customSiteUrl: string;
    filters: {
        priority: string;
        readStatus: string;
        targetAudience: string;
        dateRange: string;
    };
    showCharts: boolean;
}
export declare class DashboardComponent extends React.Component<IDashboardProps, IDashboardState> {
    private refreshTimer;
    constructor(props: IDashboardProps);
    componentDidMount(): Promise<void>;
    private handleTeamsContext;
    componentWillUnmount(): void;
    componentDidUpdate(prevProps: IDashboardProps): void;
    private isTeamsContext;
    private loadMessages;
    private handleMarkAsRead;
    private isMessageRead;
    private getPriorityColor;
    private renderTitle;
    private renderDescription;
    private renderRefreshInfo;
    private renderMessage;
    private renderPlaceholder;
    private applyFilters;
    private handleFilterChange;
    private toggleCharts;
    private getChartData;
    private renderFilters;
    private renderCharts;
    private openTeamsMessageCreator;
    private openManagerDashboard;
    private openMessageDiagnostics;
    render(): React.ReactElement<IDashboardProps>;
}
//# sourceMappingURL=DashboardComponent.d.ts.map