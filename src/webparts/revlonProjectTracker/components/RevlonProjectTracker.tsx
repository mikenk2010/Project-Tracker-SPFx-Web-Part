import * as React from "react";
import { useState, useEffect } from "react";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode
} from "@fluentui/react/lib/DetailsList";
import { ShimmeredDetailsList } from "@fluentui/react/lib/ShimmeredDetailsList";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { Panel, PanelType } from "@fluentui/react/lib/Panel";
import { TextField } from "@fluentui/react/lib/TextField";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { IProject, IProjectService } from "../../../services/IProjectService";

export interface IRevlonProjectTrackerProps {
  projectService: IProjectService;
}

export const RevlonProjectTracker: React.FunctionComponent<IRevlonProjectTrackerProps> = (props) => {
  const [projects, setProjects] = useState<IProject[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string>("");
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
  const [newProject, setNewProject] = useState<Partial<IProject>>({
    Title: "",
    ProjectStatus: "",
    ProjectManager: "",
    StartDate: "",
    EndDate: ""
  });

  useEffect(() => {
    loadProjects();
  }, []);

  const loadProjects = async () => {
    try {
      setLoading(true);
      setError("");
      const data = await props.projectService.getProjects();
      setProjects(data);
    } catch (err: any) {
      if (err.message === "LIST_NOT_FOUND") {
        setError("LIST_NOT_FOUND");
      } else {
        setError("Failed to load projects. Please try again later.");
      }
    } finally {
      setLoading(false);
    }
  };

  const handleAddProject = async () => {
    try {
      await props.projectService.addProject(newProject);
      setIsPanelOpen(false);
      setNewProject({
        Title: "",
        ProjectStatus: "",
        ProjectManager: "",
        StartDate: "",
        EndDate: ""
      });
      loadProjects();
    } catch (err: any) {
      if (err.message === "LIST_NOT_FOUND" || err.message?.includes("LIST_NOT_FOUND")) {
        setError("LIST_NOT_FOUND");
      } else {
        // Show the actual error message for debugging
        const errorMsg = err.message || "Failed to add project. Please try again.";
        setError(errorMsg);
        console.error("Error adding project:", err);
      }
    }
  };

  const columns: IColumn[] = [
    {
      key: "Title",
      name: "Project Name",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 300
    },
    {
      key: "ProjectStatus",
      name: "Status",
      fieldName: "ProjectStatus",
      minWidth: 120,
      maxWidth: 150
    },
    {
      key: "ProjectManager",
      name: "Project Manager",
      fieldName: "ProjectManager",
      minWidth: 150,
      maxWidth: 200
    },
    {
      key: "StartDate",
      name: "Start Date",
      fieldName: "StartDate",
      minWidth: 100,
      maxWidth: 120
    },
    {
      key: "EndDate",
      name: "End Date",
      fieldName: "EndDate",
      minWidth: 100,
      maxWidth: 120
    }
  ];

  const renderListNotFoundMessage = () => {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        The SharePoint list "RevlonProjects" was not found. Please create the list with the following columns:
        <ul>
          <li>Title (Single line of text)</li>
          <li>ProjectStatus (Single line of text)</li>
          <li>ProjectManager (Single line of text)</li>
          <li>StartDate (Date and Time)</li>
          <li>EndDate (Date and Time)</li>
        </ul>
      </MessageBar>
    );
  };

  return (
    <div>
      {error === "LIST_NOT_FOUND" && renderListNotFoundMessage()}
      {error && error !== "LIST_NOT_FOUND" && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError("")}>
          {error}
        </MessageBar>
      )}

      <div style={{ marginBottom: 10 }}>
        <PrimaryButton text="Add New Project" onClick={() => setIsPanelOpen(true)} />
      </div>

      <ShimmeredDetailsList
        items={projects}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.fixedColumns}
        enableShimmer={loading}
        shimmerLines={5}
      />

      <div style={{ marginTop: 20, padding: 10, textAlign: "center", color: "#666", fontSize: "12px", borderTop: "1px solid #e1dfdd" }}>
        Revlon Project Tracker v1.0.2
      </div>

      <Panel
        isOpen={isPanelOpen}
        type={PanelType.medium}
        onDismiss={() => setIsPanelOpen(false)}
        headerText="Add New Project"
      >
        <TextField
          label="Project Name"
          value={newProject.Title || ""}
          onChange={(e, value) => setNewProject({ ...newProject, Title: value || "" })}
        />
        <TextField
          label="Status"
          value={newProject.ProjectStatus || ""}
          onChange={(e, value) => setNewProject({ ...newProject, ProjectStatus: value || "" })}
        />
        <TextField
          label="Project Manager"
          value={newProject.ProjectManager || ""}
          onChange={(e, value) => setNewProject({ ...newProject, ProjectManager: value || "" })}
        />
        <TextField
          label="Start Date"
          type="date"
          value={newProject.StartDate || ""}
          onChange={(e, value) => setNewProject({ ...newProject, StartDate: value || "" })}
        />
        <TextField
          label="End Date"
          type="date"
          value={newProject.EndDate || ""}
          onChange={(e, value) => setNewProject({ ...newProject, EndDate: value || "" })}
        />
        <div style={{ marginTop: 20 }}>
          <PrimaryButton text="Save" onClick={handleAddProject} />
        </div>
      </Panel>
    </div>
  );
};

