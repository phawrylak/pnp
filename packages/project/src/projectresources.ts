import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { DraftProject } from "./projects";
import { QueueJob } from "./queuejobs";
import { CommandResult } from "./types";

/**
 * Represents a collection of resources in a published project
 */
export class PublishedProjectResourceCollection extends ProjectQueryableCollection {

    /**
    * Gets a resource from the collection with the specified GUID
    *
    * @param id The string representation of the resource GUID
    */
    public getById(id: string): PublishedProjectResource {
        const resource = new PublishedProjectResource(this);
        resource.concat(`('${id}')`);
        return resource;
    }
}

/**
 * Represents a collection of DraftProjectResource objects
 */
export class DraftProjectResourceCollection extends ProjectQueryableCollection {

    /**
    * Gets a draft project resource from the collection with the specified GUID
    *
    * @param id The string representation of the project resource GUID
    */
    public getById(id: string): DraftProjectResource {
        const resource = new DraftProjectResource(this);
        resource.concat(`('${id}')`);
        return resource;
    }

    /**
     * Adds a new project resource that is specified by the ProjectResourceCreationInformation object to the collection
     *
     * @param parameters The properties of the project resource to create
     */
    public async add(parameters: ProjectResourceCreationInformation): Promise<CommandResult<DraftProjectResource>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }

    /**
     * Adds an existing enterprise resource to the draft project resource collection
     *
     * @param resourceId The string representation of the enterprise resource GUID
     */
    public async addEnterpriseResourceById(resourceId: string): Promise<CommandResult<QueueJob>> {
        const data = await this.clone(DraftProjectResourceCollection, "AddEnterpriseResourceById").postCore({ body: jsS({ resourceId: resourceId }) });
        return { data: data, instance: this.getParent(DraftProject).queueJobs.getById(data.Id) };
    }
}

/**
 * Provides information about a project resource
 */
export abstract class ProjectResource extends ProjectQueryableInstance {
}

/**
 * Represents an enterprise resource that is published on Project Server
 */
export class PublishedProjectResource extends ProjectResource {
}

/**
 * Represents an enterprise resource in a checked-out project
 */
export class DraftProjectResource extends ProjectResource {

    /**
     * Deletes the DraftProjectResource object
     */
    public delete = this._delete;
}

/**
 * Provides property settings and methods for the creation of a project resource entity
 */
export interface ProjectResourceCreationInformation {

    /**
     * Gets or sets the project resource account
     */
    account?: string;

    /**
     * Gets or sets the email address of the project resource
     */
    email?: string;

    /**
     * Gets or sets the GUID of the project resource
     */
    id?: string;

    /**
     * Gets or sets the project resource name
     */
    name?: string;

    /**
     * Gets or sets the notes about the project resource
     */
    notes?: string;
}
