declare interface IAdaptiveCardViewerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  CardJsonUrlFieldLabel: string;
  CardJsonUrlFieldDescription: string;
}

declare interface IDashboardWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AdvancedGroupName: string;
  TitleFieldLabel: string;
  DescriptionFieldLabel: string;
  DataSourceUrlFieldLabel: string;
  DataSourceUrlFieldDescription: string;
  RefreshIntervalFieldLabel: string;
  RefreshIntervalFieldDescription: string;
  ShowRefreshButtonFieldLabel: string;
}

declare module 'AdaptiveCardViewerWebPartStrings' {
  const strings: IAdaptiveCardViewerWebPartStrings;
  export = strings;
}

declare module 'DashboardWebPartStrings' {
  const strings: IDashboardWebPartStrings;
  export = strings;
}
