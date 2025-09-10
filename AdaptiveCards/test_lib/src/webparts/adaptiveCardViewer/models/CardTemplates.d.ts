export declare const sampleCardTemplate: {
    $schema: string;
    type: string;
    version: string;
    body: ({
        type: string;
        text: string;
        weight: string;
        size: string;
        wrap?: undefined;
    } | {
        type: string;
        text: string;
        wrap: boolean;
        weight?: undefined;
        size?: undefined;
    })[];
};
/**
 * Generate Adaptive Card JSON from SharePoint message for Teams/Email distribution
 */
export declare function generateMessageCard(message: any): any;
/**
 * Generate simplified card for Teams (with enhanced read tracking)
 */
export declare function generateTeamsCard(message: any): any;
export declare const dashboardCardTemplate: {
    $schema: string;
    type: string;
    version: string;
    body: ({
        type: string;
        text: string;
        weight: string;
        size: string;
        color: string;
        facts?: undefined;
    } | {
        type: string;
        facts: {
            title: string;
            value: string;
        }[];
        text?: undefined;
        weight?: undefined;
        size?: undefined;
        color?: undefined;
    })[];
    actions: {
        type: string;
        title: string;
        data: {
            action: string;
        };
    }[];
};
export declare const cardTemplates: {
    sample: {
        $schema: string;
        type: string;
        version: string;
        body: ({
            type: string;
            text: string;
            weight: string;
            size: string;
            wrap?: undefined;
        } | {
            type: string;
            text: string;
            wrap: boolean;
            weight?: undefined;
            size?: undefined;
        })[];
    };
    dashboard: {
        $schema: string;
        type: string;
        version: string;
        body: ({
            type: string;
            text: string;
            weight: string;
            size: string;
            color: string;
            facts?: undefined;
        } | {
            type: string;
            facts: {
                title: string;
                value: string;
            }[];
            text?: undefined;
            weight?: undefined;
            size?: undefined;
            color?: undefined;
        })[];
        actions: {
            type: string;
            title: string;
            data: {
                action: string;
            };
        }[];
    };
};
//# sourceMappingURL=CardTemplates.d.ts.map