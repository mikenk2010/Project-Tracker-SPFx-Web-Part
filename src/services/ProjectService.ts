import { IProject, IProjectService } from "./IProjectService";
import { getSP } from "./pnpjsConfig";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const LIST_NAME = "RevlonProjects";

const isLocalEnvironment = (): boolean => {
  return window.location.hostname === "localhost" || 
         window.location.hostname === "127.0.0.1" ||
         window.location.hostname.includes("workbench");
};

const getMockProjects = (): IProject[] => {
  return [
    {
      Id: 1,
      Title: "Q1 Product Launch",
      ProjectStatus: "In Progress",
      ProjectManager: "John Smith",
      StartDate: "2024-01-15",
      EndDate: "2024-03-30"
    },
    {
      Id: 2,
      Title: "Website Redesign",
      ProjectStatus: "Planning",
      ProjectManager: "Sarah Johnson",
      StartDate: "2024-02-01",
      EndDate: "2024-05-15"
    },
    {
      Id: 3,
      Title: "Mobile App Development",
      ProjectStatus: "Completed",
      ProjectManager: "Mike Davis",
      StartDate: "2023-10-01",
      EndDate: "2024-01-10"
    }
  ];
};

export class ProjectService implements IProjectService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  async getProjects(): Promise<IProject[]> {
    if (isLocalEnvironment()) {
      return getMockProjects();
    }

    try {
      const sp = getSP(this.context);
      
      // First, verify the list exists
      console.log("Checking for list:", LIST_NAME);
      console.log("Current web URL:", this.context.pageContext.web.absoluteUrl);
      
      // Try to get the list first to verify it exists
      try {
        const list = await (sp as any).web.lists.getByTitle(LIST_NAME)();
        console.log("List found:", list);
      } catch (listError: any) {
        console.error("List check error:", listError);
        if (listError.status === 404) {
          throw new Error("LIST_NOT_FOUND");
        }
        throw listError;
      }
      
      const items = await (sp as any).web.lists.getByTitle(LIST_NAME).items();
      
      return items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title || "",
        ProjectStatus: item.ProjectStatus || "",
        ProjectManager: item.ProjectManager || "",
        StartDate: item.StartDate || "",
        EndDate: item.EndDate || ""
      }));
    } catch (error: any) {
      if (error.status === 404) {
        throw new Error("LIST_NOT_FOUND");
      }
      throw error;
    }
  }

  async addProject(project: Partial<IProject>): Promise<IProject> {
    if (isLocalEnvironment()) {
      const mockProject: IProject = {
        Id: Date.now(),
        Title: project.Title || "",
        ProjectStatus: project.ProjectStatus || "",
        ProjectManager: project.ProjectManager || "",
        StartDate: project.StartDate || "",
        EndDate: project.EndDate || ""
      };
      return mockProject;
    }

    try {
      const sp = getSP(this.context);
      
      // Format dates for SharePoint (convert YYYY-MM-DD to ISO string or null if empty)
      const formatDateForSharePoint = (dateString: string | undefined): string | null => {
        if (!dateString) return null;
        // If date is in YYYY-MM-DD format, convert to ISO string
        if (dateString.match(/^\d{4}-\d{2}-\d{2}$/)) {
          return new Date(dateString + 'T00:00:00').toISOString();
        }
        return dateString;
      };

      const itemData: any = {
        Title: project.Title || ""
      };

      // Only add fields if they have values
      if (project.ProjectStatus) {
        itemData.ProjectStatus = project.ProjectStatus;
      }
      if (project.ProjectManager) {
        itemData.ProjectManager = project.ProjectManager;
      }
      if (project.StartDate) {
        const formattedStartDate = formatDateForSharePoint(project.StartDate);
        if (formattedStartDate) {
          itemData.StartDate = formattedStartDate;
        }
      }
      if (project.EndDate) {
        const formattedEndDate = formatDateForSharePoint(project.EndDate);
        if (formattedEndDate) {
          itemData.EndDate = formattedEndDate;
        }
      }

      console.log("Adding project with data:", itemData);
      console.log("Current web URL:", this.context.pageContext.web.absoluteUrl);
      console.log("List name:", LIST_NAME);
      
      if (!sp) {
        throw new Error("PnPjs not properly initialized - sp is undefined");
      }
      
      // Access web property - in PnPjs v3, web is a property of sp
      const web = (sp as any).web;
      console.log("Web property:", web);
      
      if (!web) {
        console.error("SP object structure:", Object.keys(sp));
        throw new Error("PnPjs web property is undefined. SPFx context may not be properly initialized. Check browser console for details.");
      }
      
      // Verify list exists before trying to add
      try {
        const list = await web.lists.getByTitle(LIST_NAME)();
        console.log("List verified:", list.Title);
      } catch (listError: any) {
        console.error("List verification error:", listError);
        if (listError.status === 404) {
          throw new Error("LIST_NOT_FOUND");
        }
        throw new Error(`List not found: ${listError.message || 'Unknown error'}`);
      }
      
      const result = await web.lists.getByTitle(LIST_NAME).items.add(itemData);
      console.log("Project added successfully:", result.data);

      return {
        Id: result.data.Id,
        Title: result.data.Title || "",
        ProjectStatus: result.data.ProjectStatus || "",
        ProjectManager: result.data.ProjectManager || "",
        StartDate: result.data.StartDate || "",
        EndDate: result.data.EndDate || ""
      };
    } catch (error: any) {
      console.error("Error details:", error);
      if (error.status === 404) {
        throw new Error("LIST_NOT_FOUND");
      }
      // Include more detailed error information
      let errorMessage = "Failed to add project";
      if (error.message) {
        errorMessage += `: ${error.message}`;
      }
      if (error.data) {
        errorMessage += ` - ${JSON.stringify(error.data)}`;
      }
      if (error.response) {
        errorMessage += ` - Response: ${JSON.stringify(error.response)}`;
      }
      throw new Error(errorMessage);
    }
  }
}

