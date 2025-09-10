import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IDashboardWebPartProps {
    title: string;
    description: string;
    dataSourceUrl: string;
    refreshInterval: number;
    showRefreshButton: boolean;
}
export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {
    protected onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=DashboardWebPart.d.ts.map