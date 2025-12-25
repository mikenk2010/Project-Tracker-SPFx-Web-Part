export interface IProject {
  Id: number;
  Title: string;
  ProjectStatus: string;
  ProjectManager: string;
  StartDate: string;
  EndDate: string;
}

export interface IProjectService {
  getProjects(): Promise<IProject[]>;
  addProject(project: Partial<IProject>): Promise<IProject>;
}

