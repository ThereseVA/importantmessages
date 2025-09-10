import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IManager {
  Id: number;
  Title: string;
  ManagerEmail: {
    EMail: string;
    Title: string;
  };
  ManagerDisplayName: string;
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
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this.listName}')/items?$select=Id,Title,ManagerEmail/EMail,ManagerEmail/Title,ManagerDisplayName,Department,ManagerLevel,IsActive,StartDate,EndDate,Notes&$expand=ManagerEmail&$filter=IsActive eq true&$orderby=ManagerDisplayName`;

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      return data.value as IManager[];
    } catch (error) {
      console.error('Error fetching managers from SharePoint list:', error);
      throw error;
    }
  }

  /**
   * Check if a specific user is a manager
   */
  public async isUserManager(userEmail: string): Promise<boolean> {
    try {
      const managers = await this.getActiveManagers();
      return managers.some(manager => 
        manager.ManagerEmail?.EMail?.toLowerCase() === userEmail.toLowerCase()
      );
    } catch (error) {
      console.error('Error checking if user is manager:', error);
      return false;
    }
  }

  /**
   * Get manager details for a specific user
   */
  public async getManagerDetails(userEmail: string): Promise<IManager | null> {
    try {
      const managers = await this.getActiveManagers();
      return managers.find(manager => 
        manager.ManagerEmail?.EMail?.toLowerCase() === userEmail.toLowerCase()
      ) || null;
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
        Title: manager.Title || manager.ManagerDisplayName,
        ManagerDisplayName: manager.ManagerDisplayName,
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
        ManagerDisplayName: updates.ManagerDisplayName,
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
