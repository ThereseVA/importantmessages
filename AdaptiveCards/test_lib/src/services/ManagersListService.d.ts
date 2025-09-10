import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IManager {
    Id: number;
    Title: string;
    ManagersEmail: {
        EMail: string;
        Title: string;
    };
    ManagersDisplayName: string;
    Department: string;
    ManagerLevel: string;
    IsActive: boolean;
    StartDate: string;
    EndDate: string;
    Notes: string;
}
export declare class ManagersListService {
    private context;
    private listName;
    constructor(context: WebPartContext);
    /**
     * Get all active managers from the Managers list
     */
    getActiveManagers(): Promise<IManager[]>;
    /**
     * Check if a specific user is a manager
     */
    isUserManager(userEmail: string): Promise<boolean>;
    /**
     * Get manager details for a specific user
     */
    getManagerDetails(userEmail: string): Promise<IManager | null>;
    /**
     * Get managers by department
     */
    getManagersByDepartment(department: string): Promise<IManager[]>;
    /**
     * Get managers by level
     */
    getManagersByLevel(level: string): Promise<IManager[]>;
    /**
     * Add a new manager to the list
     */
    addManager(manager: Partial<IManager>): Promise<boolean>;
    /**
     * Update an existing manager
     */
    updateManager(managerId: number, updates: Partial<IManager>): Promise<boolean>;
    /**
     * Deactivate a manager (set IsActive to false)
     */
    deactivateManager(managerId: number, endDate?: string): Promise<boolean>;
    /**
     * Check if the current user has permission to manage the Managers list
     */
    canManageManagersList(): Promise<boolean>;
}
//# sourceMappingURL=ManagersListService.d.ts.map