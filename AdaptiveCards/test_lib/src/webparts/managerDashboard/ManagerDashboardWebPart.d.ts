import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IManagerDashboardWebPartProps {
    title: string;
}
export default class ManagerDashboardWebPart extends BaseClientSideWebPart<IManagerDashboardWebPartProps> {
    private _isDarkTheme;
    private _environmentMessage;
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected onAfterPropertyPaneChangesApplied(): void;
}
//# sourceMappingURL=ManagerDashboardWebPart.d.ts.map