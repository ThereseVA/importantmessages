import * as React from 'react';
import { useState, useEffect } from 'react';
import { Dropdown, IDropdownOption, MessageBar, MessageBarType, DefaultButton, PrimaryButton } from 'office-ui-fabric-react';
import { enhancedDataService } from '../../../services/EnhancedDataService';

export interface ISiteSelectorProps {
  onSiteSelected: (siteUrl: string, siteName: string) => void;
  currentSite?: string;
}

export const SiteSelector: React.FunctionComponent<ISiteSelectorProps> = (props) => {
  const [availableSites, setAvailableSites] = useState<IDropdownOption[]>([]);
  const [selectedSite, setSelectedSite] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');

  // Predefined sites for Gustav Kliniken - only include sites where Important Messages list exists
  const predefinedSites: IDropdownOption[] = [
    // Removed root site since Important Messages list doesn't exist there
    // { key: 'https://gustafkliniken.sharepoint.com/', text: 'Gustafkliniken (Root Site)', data: { name: 'GustafklinikenRoot' } },
    { key: 'https://gustafkliniken.sharepoint.com/sites/Gustafkliniken', text: 'Gustafkliniken (Main Site)', data: { name: 'Gustafkliniken' } }
  ];

  useEffect(() => {
    // Set default sites
    setAvailableSites(predefinedSites);
    
    // Set current site if provided
    if (props.currentSite) {
      setSelectedSite(props.currentSite);
    } else {
      // Always default to the main site where Important Messages list exists
      setSelectedSite('https://gustafkliniken.sharepoint.com/sites/Gustafkliniken');
    }
  }, [props.currentSite]);

  const handleSiteChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      setSelectedSite(option.key as string);
    }
  };

  const handleSiteSelection = (): void => {
    if (selectedSite) {
      const site = availableSites.find(s => s.key === selectedSite);
      if (site) {
        props.onSiteSelected(selectedSite, site.data?.name || site.text);
      }
    }
  };

  const testSiteConnection = async (): Promise<void> => {
    if (!selectedSite) return;

    setIsLoading(true);
    setError('');

    try {
      enhancedDataService.setSharePointSiteUrl(selectedSite);
      
      // Test connection by trying to get site info
      const response = await fetch(`${selectedSite}/_api/web`, {
        headers: {
          'Accept': 'application/json;odata=nometadata'
        }
      });

      if (response.ok) {
        setError('');
        // Auto-select if connection test passes
        handleSiteSelection();
      } else {
        setError('Cannot connect to this site. Please check permissions.');
      }
    } catch (err) {
      setError('Failed to connect to site. Please verify the URL and your access permissions.');
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div style={{ padding: '20px', border: '1px solid #edebe9', borderRadius: '4px', marginBottom: '20px' }}>
      <h3>🏢 Select SharePoint Site</h3>
      <p>Choose which SharePoint site you want to manage messages for:</p>
      
      <Dropdown
        label="SharePoint Site"
        options={availableSites}
        selectedKey={selectedSite}
        onChange={handleSiteChange}
        placeholder="Select a site..."
        style={{ marginBottom: '15px', maxWidth: '400px' }}
      />

      {error && (
        <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: '15px' }}>
          {error}
        </MessageBar>
      )}

      <div style={{ display: 'flex', gap: '10px' }}>
        <PrimaryButton 
          text="Select Site" 
          onClick={handleSiteSelection}
          disabled={!selectedSite || isLoading}
        />
        
        <DefaultButton 
          text="Test Connection" 
          onClick={testSiteConnection}
          disabled={!selectedSite || isLoading}
        />
      </div>

      {selectedSite && (
        <div style={{ marginTop: '15px', padding: '10px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
          <strong>Selected Site:</strong> {selectedSite}
          <br />
          <small>Click "Select Site" to start managing messages for this site.</small>
        </div>
      )}
    </div>
  );
};
