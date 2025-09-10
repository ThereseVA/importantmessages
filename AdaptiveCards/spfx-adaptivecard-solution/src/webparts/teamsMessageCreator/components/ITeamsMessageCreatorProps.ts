import { WebPartContext } from '@microsoft/sp-webpart-base';
import { EnhancedDataService } from '../../../services/EnhancedDataService';

export interface ITeamsMessageCreatorProps {
  title: string;
  context: WebPartContext;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  dataService: EnhancedDataService;
}
