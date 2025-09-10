import * as React from 'react';
import { ITeamsMessageCreatorProps } from './ITeamsMessageCreatorProps';
interface ITeamsMessageCreatorComponentState {
    isManager: boolean;
    loading: boolean;
    error: string | null;
}
export declare class TeamsMessageCreatorComponent extends React.Component<ITeamsMessageCreatorProps, ITeamsMessageCreatorComponentState> {
    private dataService;
    constructor(props: ITeamsMessageCreatorProps);
    componentDidMount(): Promise<void>;
    render(): React.ReactElement<ITeamsMessageCreatorProps>;
}
export {};
//# sourceMappingURL=TeamsMessageCreatorComponent.d.ts.map