import * as React from 'react';
import { IAdaptiveCardProps } from './IAdaptiveCardProps';
export interface IAdaptiveCardState {
    cardData: any;
    loading: boolean;
    error: string | null;
}
export declare class AdaptiveCardComponent extends React.Component<IAdaptiveCardProps, IAdaptiveCardState> {
    private cardContainer;
    constructor(props: IAdaptiveCardProps);
    componentDidMount(): Promise<void>;
    componentDidUpdate(prevProps: IAdaptiveCardProps): void;
    private loadCard;
    private loadAssetCard;
    private renderAdaptiveCard;
    private handleSubmitAction;
    private showSuccessMessage;
    private showErrorMessage;
    private getDefaultCard;
    private renderPlaceholder;
    private renderTitle;
    private renderComponent;
    render(): React.ReactElement<IAdaptiveCardProps>;
}
//# sourceMappingURL=AdaptiveCardComponent.d.ts.map