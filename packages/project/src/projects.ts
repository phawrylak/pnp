import { jsS, TypedHash } from "@pnp/common";
import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { DraftAssignmentCollection, PublishedAssignmentCollection } from "./assignments";
import { Calendar } from "./calendars";
import { CustomFieldCollection } from "./customfields";
import { EnterpriseProjectType } from "./enterpriseprojecttypes";
import { Phase } from "./phases";
import { DraftProjectResourceCollection, PublishedProjectResourceCollection } from "./projectresources";
import { QueueJob, QueueJobCollection } from "./queuejobs";
import { Stage } from "./stages";
import { DraftTaskLinkCollection, PublishedTaskLinkCollection } from "./tasklinks";
import {
    DraftTaskCollection,
    ProjectSummaryTask,
    PublishedTaskCollection,
} from "./tasks";
import { User } from "./users";

/**
 * Represents a collection of PublishedProject objects
 */
@defaultPath("_api/ProjectServer/Projects")
export class ProjectCollection extends ProjectQueryableCollection {

    /**
    * Gets a project from the collection with the specified GUID
    *
    * @param id The string representation of the project GUID
    */
    public getById(id: string): PublishedProject {
        const project = new PublishedProject(this);
        project.concat(`('${id}')`);
        return project;
    }

    /**
     * Adds the project that is specified by the ProjectCreationInformation object to the collection
     *
     * @param parameters The properties of the project to create
     */
    public async add(parameters: ProjectCreationInformation): Promise<CommandResult<PublishedProject>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Contains the common properties for draft projects and published projects
 */
export abstract class Project extends ProjectQueryableInstance {

    /**
     * Gets the enterprise resource who has the project checked out
     */
    public get checkedOutBy(): User {
        return new User(this, "CheckedOutBy");
    }

    /**
     * Gets the collection of project custom fields that have values set for the project
     */
    public get customFields(): CustomFieldCollection {
        return new CustomFieldCollection(this, "CustomFields");
    }

    /**
     * Gets the enterprise project type (EPT) for the project
     */
    public get enterpriseProjectType(): EnterpriseProjectType {
        return new EnterpriseProjectType(this, "EnterpriseProjectType");
    }

    /**
     * Gets the current workflow phase of the project
     */
    public get phase(): Phase {
        return new Phase(this, "Phase");
    }

    /**
     * Gets the hidden project summary task
     */
    public get projectSummaryTask(): ProjectSummaryTask {
        return new ProjectSummaryTask(this, "ProjectSummaryTask");
    }

    /**
     * Gets the collection of Project Server Queue Service jobs that are associated with the project
     */
    public get queueJobs(): QueueJobCollection {
        return new QueueJobCollection(this, "QueueJobs");
    }

    /**
     * Gets the current workflow stage of the project
     */
    public get stage(): Stage {
        return new Stage(this, "Stage");
    }
}

/**
 * Represents a project that is published on Project Server
 */
export class PublishedProject extends Project {

    /**
     * Gets the collection of assignments in the project
     */
    public get assignments(): PublishedAssignmentCollection {
        return new PublishedAssignmentCollection(this, "Assignments");
    }

    /**
     * Gets the project calendar
     */
    public get calendar(): Calendar {
        return new Calendar(this, "Calendar");
    }

    /**
     * Gets a DraftProject object if it is not already checked out
     */
    public get draft(): DraftProject {
        return new DraftProject(this, "Draft");
    }

    /**
     * Gets a PublishedProject object that includes custom fields
     */
    public get includeCustomFields(): PublishedProject {
        return new PublishedProject(this, "IncludeCustomFields");
    }

    /**
     * Gets the owner of the project
     */
    public get owner(): User {
        return new User(this, "Owner");
    }

    /**
     * Gets the collection of resources for a project
     */
    public get projectResources(): PublishedProjectResourceCollection {
        return new PublishedProjectResourceCollection(this, "ProjectResources");
    }

    /**
     * Gets the collection of task links for the project
     */
    public get taskLinks(): PublishedTaskLinkCollection {
        return new PublishedTaskLinkCollection(this, "TaskLinks");
    }

    /**
     * Gets the collection of tasks for the project
     */
    public get tasks(): PublishedTaskCollection {
        return new PublishedTaskCollection(this, "Tasks");
    }

    /**
     * Checks out the draft version of the project
     */
    public async checkOut(): Promise<CommandResult<DraftProject>> {
        const data = await this.clone(PublishedProject, "CheckOut").postCore();
        return { data: data, instance: this.draft };
    }

    /**
     * Deletes the PublishedProject object
     */
    public async delete(): Promise<CommandResult<QueueJob>> {
        const data = await this.postCore({
            headers: {
                "X-HTTP-Method": "DELETE",
            },
        });
        return { data: data, instance: this.queueJobs.getById(data.Id) };
    }

    /**
     * Causes the workflow to run
     */
    public submitToWorkflow(): Promise<void> {
        return this.clone(PublishedProject, "SubmitToWorkflow").postCore();
    }
}

/**
 * Represents the draft version of a project, which is a project that is checked out
 */
export class DraftProject extends Project {

    /**
     * Gets the collection of assignments for a project
     */
    public get assignments(): DraftAssignmentCollection {
        return new DraftAssignmentCollection(this, "Assignments");
    }

    /**
     * Gets a Project Server calendar
     */
    public get calendar(): Calendar {
        return new Calendar(this, "Calendar");
    }

    /**
     * Gets a DraftProject object that includes custom fields
     */
    public get includeCustomFields(): DraftProject {
        return new DraftProject(this, "IncludeCustomFields");
    }

    /**
     * Gets the project owner
     */
    public get owner(): User {
        return new User(this, "Owner");
    }

    /**
     * Gets the collection of resources for a project
     */
    public get projectResources(): DraftProjectResourceCollection {
        return new DraftProjectResourceCollection(this, "ProjectResources");
    }

    /**
     * Gets the collection of draft task link objects for the project
     */
    public get taskLinks(): DraftTaskLinkCollection {
        return new DraftTaskLinkCollection(this, "TaskLinks");
    }

    /**
     * Gets the collection of task objects for the project
     */
    public get tasks(): DraftTaskCollection {
        return new DraftTaskCollection(this, "Tasks");
    }

    /**
    * Updates this draft project with the supplied properties
    *
    * @param properties A plain object of property names and values to update the draft project
    */
    public update = this._update<CommandResult<QueueJob>, TypedHash<any>, any>(
        "PS.DraftProject",
        data => ({ data: data, instance: this.queueJobs.getById(data.Id) }));

    /**
     * Queues a check-in job for a draft project if it is still checked out
     *
     * @param force True if the administrator or project owner forces check in of a project; otherwise, false
     */
    public async checkIn(force: boolean): Promise<CommandResult<QueueJob>> {
        const data = await this.clone(DraftProject, "CheckIn").postCore({ body: jsS({ force: force }) });
        return { data: data, instance: this.queueJobs.getById(data.Id) };
    }

    /**
     * Queues a publish job to get the changes from the draft project back to the published version
     *
     * @param checkIn Boolean that indicates whether the project should be checked in after it is published
     */
    public async publish(checkIn: boolean): Promise<CommandResult<QueueJob>> {
        const data = await this.clone(DraftProject, "Publish").postCore({ body: jsS({ checkIn: checkIn }) });
        return { data: data, instance: this.queueJobs.getById(data.Id) };
    }
}

/**
 * Contains the properties that can be set when creating a project
 */
export interface ProjectCreationInformation {

    /**
     * Gets or sets the description of the project
     */
    description?: string;

    /**
     * Gets or sets the GUID of the enterprise project type (EPT)
     */
    enterpriseProjectTypeId?: string;

    /**
     * Gets or sets the GUID of the project
     */
    id?: string;

    /**
     * Gets or sets the name of the project
     */
    name: string;

    /**
     * Gets or sets the start date of the project
     */
    start?: Date;
}
