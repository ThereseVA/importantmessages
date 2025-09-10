import * as React from 'react';
import { IManagerDashboardProps } from './IManagerDashboardProps';
interface IManagerDashboardComponentState {
    isManager: boolean;
    loading: boolean;
    error: string | null;
}
export declare class ManagerDashboardComponent extends React.Component<IManagerDashboardProps, IManagerDashboardComponentState> {
    constructor(props: IManagerDashboardProps);
    componentDidMount(): Promise<void>;
    render(): React.ReactElement<IManagerDashboardProps>;
}
export {};
//# sourceMappingURL=ManagerDashboardComponent.d.ts.map