import { WebPartContext } from '@microsoft/sp-webpart-base';
import { EnhancedDataService } from '../../../services/EnhancedDataService';

export interface IManagerDashboardProps {
  title: string;
  context: WebPartContext;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  dataService: EnhancedDataService;
}
