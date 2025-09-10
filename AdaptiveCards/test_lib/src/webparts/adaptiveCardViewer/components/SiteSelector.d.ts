import * as React from 'react';
export interface ISiteSelectorProps {
    onSiteSelected: (siteUrl: string, siteName: string) => void;
    currentSite?: string;
}
export declare const SiteSelector: React.FunctionComponent<ISiteSelectorProps>;
//# sourceMappingURL=SiteSelector.d.ts.map