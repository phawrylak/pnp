import { jsS, TypedHash } from "@pnp/common";
import {
    defaultPath,
    SharePointQueryableCollection,
    SharePointQueryableInstance,
} from "@pnp/sp";

import { QueueJob, QueueJobs } from "./queueJobs";

/**
 * Represents a collection of PublishedProject objects
 */
@defaultPath("_api/ProjectServer/Projects")
export class Projects extends SharePointQueryableCollection {

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
    public async add(parameters: ProjectCreationInformation): Promise<PublishedProjectResult> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, project: this.getById(data.Id) };
    }
}

/**
 * Contains the common properties for draft projects and published projects
 */
export class Project extends SharePointQueryableInstance {

    /**
     * Gets the collection of Project Server Queue Service jobs that are associated with the project
     */
    public get queueJobs(): QueueJobs {
        return new QueueJobs(this, "QueueJobs");
    }
}

/**
 * Represents a project that is published on Project Server
 */
export class PublishedProject extends Project {

    /**
     * Deletes this project
     *
     */
    public delete = this._delete;

    /**
     * Gets a DraftProject object if it is not already checked out
     */
    public get draft(): DraftProject {
        return new DraftProject(this, "Draft");
    }

    /**
     * Checks out the draft version of the project
     */
    public async checkOut(): Promise<DraftProjectResult> {
        const data = await this.clone(PublishedProject, "CheckOut").postCore();
        return { data: data, draft: this.draft };
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
    * Updates this draft project with the supplied properties
    *
    * @param properties A plain object of property names and values to update the draft project
    */
    public update = this._update<QueueJobResult, TypedHash<any>, any>(
        "PS.DraftProject",
        data => ({ data: data, queueJob: this.queueJobs.getById(data.Id) }));

    /**
     * Queues a check-in job for a draft project if it is still checked out
     *
     * @param force True if the administrator or project owner forces check in of a project; otherwise, false
     */
    public async checkIn(force: boolean): Promise<QueueJobResult> {
        const data = await this.clone(DraftProject, "CheckIn").postCore({ body: jsS({ force: force }) });
        return { data: data, queueJob: this.queueJobs.getById(data.Id) };
    }

    /**
     * Queues a publish job to get the changes from the draft project back to the published version
     *
     * @param checkIn Boolean that indicates whether the project should be checked in after it is published
     */
    public async publish(checkIn: boolean): Promise<QueueJobResult> {
        const data = await this.clone(DraftProject, "Publish").postCore({ body: jsS({ checkIn: checkIn }) });
        return { data: data, queueJob: this.queueJobs.getById(data.Id) };
    }
}

/**
 * Contains the properties that can be set when creating a project
 */
export interface ProjectCreationInformation {

    /**
     * Gets or sets the description of the project
     */
    Description?: string;

    /**
     * Gets or sets the GUID of the enterprise project type (EPT)
     */
    EnterpriseProjectTypeId?: string;

    /**
     * Gets or sets the GUID of the project
     */
    Id?: string;

    /**
     * Gets or sets the name of the project
     */
    Name: string;

    /**
     * Gets or sets the start date of the project
     */
    Start?: Date;
}

export interface PublishedProjectResult {
    data: any;
    project: PublishedProject;
}

export interface DraftProjectResult {
    data: any;
    draft: DraftProject;
}

export interface QueueJobResult {
    data: any;
    queueJob: QueueJob;
}
