import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
export interface IAdaptiveCardProps {
    cardJsonUrl: string;
    title: string;
    enableActions: boolean;
    context: WebPartContext;
    displayMode: DisplayMode;
    updateProperty: (value: string) => void;
}
//# sourceMappingURL=IAdaptiveCardProps.d.ts.map