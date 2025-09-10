import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IDashboardProps {
  title: string;
  description: string;
  dataSourceUrl: string;
  refreshInterval: number;
  showRefreshButton: boolean;
  context: WebPartContext;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
