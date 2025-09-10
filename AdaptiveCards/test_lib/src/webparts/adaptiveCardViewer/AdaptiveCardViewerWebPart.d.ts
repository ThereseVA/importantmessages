import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IAdaptiveCardViewerWebPartProps {
    cardJsonUrl: string;
    title: string;
    enableActions: boolean;
    cardSource: string;
}
export default class AdaptiveCardViewerWebPart extends BaseClientSideWebPart<IAdaptiveCardViewerWebPartProps> {
    protected onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=AdaptiveCardViewerWebPart.d.ts.map