import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IManager {
  Id: number;
  Title: string;
  ManagersEmail: {
    EMail: string;
    Title: string;
  };
  ManagersDisplayName: string;
  EmailAdress: string; // Add the text field for email
  Department: string;
  ManagerLevel: string;
  IsActive: boolean;
  StartDate: string;
  EndDate: string;
  Notes: string;
}

export class ManagersListService {
  private context: WebPartContext;
  private listName: string = "Managers";

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Get all active managers from the Managers list
   */
  public async getActiveManagers(): Promise<IManager[]> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      console.log('🚀🚀🚀 NUCLEAR CACHE BREAK v3.0.001 - TIMESTAMP:', Date.now());
      console.log('💥💥💥 MANAGERS LIST SERVICE - NUCLEAR v3.0.001 - ABSOLUTE CACHE DESTRUCTION 💥💥💥');
      console.log('🔍 ManagersListService.getActiveManagers() - Starting...');
      console.log('🔍 Site URL:', siteUrl);
      console.log('🔍 List Name:', this.listName);
      console.log('🔍 🔍 🔍 ENHANCED DEBUGGING - CHECKING IF LIST EXISTS 🔍 🔍 🔍');
      
      // First, let's check if the list exists
      const listCheckUrl = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')`;
      console.log('🔍 Checking if Managers list exists:', listCheckUrl);
      
      try {
        const listCheckResponse = await this.context.spHttpClient.get(
          listCheckUrl,
          SPHttpClient.configurations.v1
        );
        console.log('🔍 List check response status:', listCheckResponse.status);
        
        if (listCheckResponse.ok) {
          const listData = await listCheckResponse.json();
          console.log('🔍 Managers list found:', listData.Title);
          console.log('🔍 List ID:', listData.Id);
          console.log('🔍 Item count:', listData.ItemCount);
        } else {
          console.error('🔍 Managers list not found, status:', listCheckResponse.status);
          const errorText = await listCheckResponse.text();
          console.error('🔍 List check error:', errorText);
          
          // Try to find all lists to see what's available
          const allListsUrl = `${siteUrl}/_api/web/lists?$select=Title,Id,ItemCount`;
          console.log('🔍 Getting all lists to see what exists...');
          const allListsResponse = await this.context.spHttpClient.get(
            allListsUrl,
            SPHttpClient.configurations.v1
          );
          
          if (allListsResponse.ok) {
            const allListsData = await allListsResponse.json();
            console.log('🔍 All available lists:', allListsData.value.map((list: any) => ({
              Title: list.Title,
              Id: list.Id,
              ItemCount: list.ItemCount
            })));
            
            // Look for lists that might contain "manager" in the name
            const managerLists = allListsData.value.filter((list: any) => 
              list.Title.toLowerCase().includes('manager') || 
              list.Title.toLowerCase().includes('manage')
            );
            console.log('🔍 Lists containing "manager" or "manage":', managerLists);
          }
          
          return [];
        }
      } catch (listCheckError) {
        console.error('🔍 Error checking if list exists:', listCheckError);
        return [];
      }

      // Now try to get the actual data with maximum cache busting
      const timestamp = Date.now();
      const random = Math.random().toString(36).substring(7);
      const itemsUrl = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Id,Title,ManagersEmail,ManagersDisplayName,EmailAdress,Department,ManagerLevel,IsActive,StartDate,EndDate,Notes&$top=5000&_t=${timestamp}&_r=${random}&cachebuster=${Date.now()}`;
      
      console.log('🔍 Fetching managers with URL:', itemsUrl);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        itemsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
          }
        }
      );

      console.log('🔍 Managers fetch response status:', response.status);

      if (!response.ok) {
        console.error('🔍 Failed to fetch managers, status:', response.status);
        const errorText = await response.text();
        console.error('🔍 Error details:', errorText);
        throw new Error(`Failed to fetch managers: ${response.status} - ${errorText}`);
      }

      const data = await response.json();
      console.log('🔍 Raw managers data received:', data);
      console.log('🔍 Number of items returned:', data.value?.length || 0);
      
      if (data.value && data.value.length > 0) {
        data.value.forEach((item: any, index: number) => {
          console.log(`🔍 Manager item ${index + 1}:`, {
            Id: item.Id,
            Title: item.Title,
            ManagersEmail: item.ManagersEmail,
            ManagersDisplayName: item.ManagersDisplayName,
            EmailAdress: item.EmailAdress,
            IsActive: item.IsActive
          });
        });
      }

      const managers: IManager[] = data.value || [];
      const activeManagers = managers.filter(manager => manager.IsActive !== false);
      
      console.log('🔍 Total managers:', managers.length);
      console.log('🔍 Active managers:', activeManagers.length);
      
      return activeManagers;

    } catch (error) {
      console.error('🔍 Error fetching managers from SharePoint list:', error);
      throw error;
    }
  }

  /**
   * Check if a specific user is a manager
   */
  public async isUserManager(userEmail: string): Promise<boolean> {
    console.log('🚀🚀🚀 NUCLEAR CACHE BREAK v3.0.001 - isUserManager() called');
    console.log('🔥 Input user email:', userEmail);
    console.log('🔥 Browser timestamp:', new Date().toISOString());
    console.log('🔥 SharePoint site URL:', this.context.pageContext.web.absoluteUrl);
    
    try {
      const managers = await this.getActiveManagers();
      
      console.log('🔍 NUCLEAR: Total active managers found:', managers?.length || 0);
      
      if (!managers || managers.length === 0) {
        console.error('❌ NUCLEAR: No managers found in list! This is the core issue.');
        
        // 🚀 NUCLEAR: Let's manually check what's in the list with raw REST
        console.log('🔍 NUCLEAR: Attempting manual REST call to debug list contents...');
        const siteUrl = this.context.pageContext.web.absoluteUrl;
        const rawUrl = `${siteUrl}/_api/web/lists/getbytitle('Managers')/items?$select=*&$top=5000`;
        
        try {
          const rawResponse = await this.context.spHttpClient.get(rawUrl, SPHttpClient.configurations.v1);
          console.log('🔍 Raw REST response status:', rawResponse.status);
          const rawData = await rawResponse.json();
          console.log('🔍 Raw REST data:', rawData);
          
          if (rawData.value) {
            console.log('🔍 Raw list items count:', rawData.value.length);
            rawData.value.forEach((item: any, index: number) => {
              console.log(`🔍 Raw item ${index + 1}:`, item);
            });
          }
        } catch (rawError) {
          console.error('❌ Raw REST call failed:', rawError);
        }
        
        return false;
      }
      
      console.log('🔍 NUCLEAR: Starting enhanced manager comparison...');
      
      let foundMatch = false;
      const userEmailLower = userEmail.toLowerCase().trim();
      
      managers.forEach((manager, index) => {
        console.log(`🔍 NUCLEAR: Manager ${index + 1} detailed analysis:`, {
          Id: manager.Id,
          Title: manager.Title,
          ManagersDisplayName: manager.ManagersDisplayName,
          ManagersEmail: manager.ManagersEmail,
          ManagersEmailType: typeof manager.ManagersEmail,
          ManagersEmailStringified: JSON.stringify(manager.ManagersEmail),
          EmailAdress: manager.EmailAdress,
          EmailAdressType: typeof manager.EmailAdress,
          AllProperties: Object.keys(manager)
        });
        
        // Extract all possible email values
        const possibleEmails: string[] = [];
        
        if (manager.ManagersEmail?.EMail && typeof manager.ManagersEmail.EMail === 'string') {
          possibleEmails.push(manager.ManagersEmail.EMail.toLowerCase().trim());
        }
        if (manager.EmailAdress && typeof manager.EmailAdress === 'string') {
          possibleEmails.push(manager.EmailAdress.toLowerCase().trim());
        }
        if (manager.Title && typeof manager.Title === 'string') {
          possibleEmails.push(manager.Title.toLowerCase().trim());
        }
        if (manager.ManagersDisplayName && typeof manager.ManagersDisplayName === 'string') {
          possibleEmails.push(manager.ManagersDisplayName.toLowerCase().trim());
        }
        
        console.log(`🔍 NUCLEAR: Manager ${index + 1} possible emails:`, possibleEmails);
        
        for (const email of possibleEmails) {
          console.log(`🔍 NUCLEAR: Comparing "${email}" vs "${userEmailLower}"`);
          
          if (email === userEmailLower || 
              email.includes(userEmailLower) || 
              userEmailLower.includes(email)) {
            console.log('✅ NUCLEAR: MATCH FOUND!');
            foundMatch = true;
            break;
          }
        }
      });
      
      console.log('🔍 NUCLEAR: Final result - Is user a manager?', foundMatch);
      console.log('🔍 NUCLEAR: User email checked:', userEmailLower);
      console.log('🔍 NUCLEAR: Total managers processed:', managers.length);
      
      return foundMatch;
      
    } catch (error) {
      console.error('❌ NUCLEAR: Error in isUserManager:', error);
      console.error('❌ Error stack:', error.stack);
      return false;
    }
  }

  /**
   * Get manager details for a specific user
   */
  public async getManagerDetails(userEmail: string): Promise<IManager | null> {
    try {
      const managers = await this.getActiveManagers();
      
      return managers.find(manager => {
        const emailFromEMail = manager.ManagersEmail?.EMail?.toLowerCase();
        const emailFromEmailAdress = manager.EmailAdress?.toLowerCase();
        const userEmailLower = userEmail.toLowerCase();
        
        return emailFromEMail === userEmailLower || 
               emailFromEmailAdress === userEmailLower;
      }) || null;
      
    } catch (error) {
      console.error('Error getting manager details:', error);
      return null;
    }
  }

  /**
   * Get managers by department
   */
  public async getManagersByDepartment(department: string): Promise<IManager[]> {
    try {
      const managers = await this.getActiveManagers();
      return managers.filter(manager => 
        manager.Department?.toLowerCase() === department.toLowerCase()
      );
    } catch (error) {
      console.error('Error getting managers by department:', error);
      return [];
    }
  }

  /**
   * Get managers by level
   */
  public async getManagersByLevel(level: string): Promise<IManager[]> {
    try {
      const managers = await this.getActiveManagers();
      return managers.filter(manager => 
        manager.ManagerLevel?.toLowerCase() === level.toLowerCase()
      );
    } catch (error) {
      console.error('Error getting managers by level:', error);
      return [];
    }
  }

  /**
   * Add a new manager to the list
   */
  public async addManager(manager: Partial<IManager>): Promise<boolean> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items`;

      const body = JSON.stringify({
        Title: manager.Title || manager.ManagersDisplayName,
        ManagersDisplayName: manager.ManagersDisplayName,
        Department: manager.Department,
        ManagerLevel: manager.ManagerLevel,
        IsActive: manager.IsActive !== undefined ? manager.IsActive : true,
        StartDate: manager.StartDate,
        EndDate: manager.EndDate,
        Notes: manager.Notes
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        listUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: body
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error adding manager to SharePoint list:', error);
      return false;
    }
  }

  /**
   * Update an existing manager
   */
  public async updateManager(managerId: number, updates: Partial<IManager>): Promise<boolean> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items(${managerId})`;

      // Get the item's etag first
      const getResponse = await this.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );
      
      if (!getResponse.ok) {
        throw new Error('Manager not found');
      }

      const etag = getResponse.headers.get('ETag');

      const body = JSON.stringify({
        Title: updates.Title,
        ManagersDisplayName: updates.ManagersDisplayName,
        Department: updates.Department,
        ManagerLevel: updates.ManagerLevel,
        IsActive: updates.IsActive,
        StartDate: updates.StartDate,
        EndDate: updates.EndDate,
        Notes: updates.Notes
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        listUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'X-HTTP-Method': 'MERGE',
            'If-Match': etag || '*'
          },
          body: body
        }
      );

      return response.ok;
    } catch (error) {
      console.error('Error updating manager in SharePoint list:', error);
      return false;
    }
  }

  /**
   * Deactivate a manager (set IsActive to false)
   */
  public async deactivateManager(managerId: number, endDate?: string): Promise<boolean> {
    return await this.updateManager(managerId, {
      IsActive: false,
      EndDate: endDate || new Date().toISOString()
    });
  }

  /**
   * Check if the current user has permission to manage the Managers list
   */
  public async canManageManagersList(): Promise<boolean> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/effectivebasepermissions`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return false;
      }

      const data = await response.json();
      // Check for AddListItems permission (value 2)
      return (data.High & 0) !== 0 || (data.Low & 2) !== 0;
    } catch (error) {
      console.error('Error checking permissions for Managers list:', error);
      return false;
    }
  }
}
