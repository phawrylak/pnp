import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { CustomFieldCollection } from "./customfields";
import { DraftProjectResource, PublishedProjectResource } from "./projectresources";
import { DraftTask, PublishedTask } from "./tasks";
import { User } from "./users";

/**
 * Represents a collection of published assignments
 */
export class PublishedAssignmentCollection extends ProjectQueryableCollection {

    /**
    * Gets a published assignment from the collection with the specified GUID
    *
    * @param id The string representation of the assignment GUID
    */
    public getById(id: string): PublishedAssignment {
        const assignment = new PublishedAssignment(this);
        assignment.concat(`('${id}')`);
        return assignment;
    }
}

/**
 * Represents a collection of DraftAssignment objects
 */
export class DraftAssignmentCollection extends ProjectQueryableCollection {

    /**
    * Gets a draft assignment from the collection with the specified GUID
    *
    * @param id The string representation of the assignment GUID
    */
    public getById(id: string): DraftAssignment {
        const assignment = new DraftAssignment(this);
        assignment.concat(`('${id}')`);
        return assignment;
    }

    /**
     * Adds the draft assignment that is specified by the AssignmentCreationInformation object to the collection
     *
     * @param parameters The properties of the draft assignment to create
     */
    public async add(parameters: AssignmentCreationInformation): Promise<CommandResult<DraftAssignment>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Contains the common properties for draft assignments and published assignments
 */
export abstract class Assignment extends ProjectQueryableInstance {

    /**
     * Gets the collection of custom fields for the assignment
     */
    public get customFields(): CustomFieldCollection {
        return new CustomFieldCollection(this, "CustomFields");
    }
}

/**
 * Represents the assignment that is in a published project
 */
export class PublishedAssignment extends Assignment {

    /**
     * Gets the user who is responsible for entering status for the current assignment
     */
    public get owner(): User {
        return new User(this, "Owner");
    }

    /**
     * Gets the parent assignment link
     */
    public get parent(): PublishedAssignment {
        return new PublishedAssignment(this, "Parent");
    }

    /**
     * Gets the resource that is associated with the assignment
     */
    public get resource(): PublishedProjectResource {
        return new PublishedProjectResource(this, "Resource");
    }

    /**
     * Gets the task to which the assignment belongs
     */
    public get task(): PublishedTask {
        return new PublishedTask(this, "Task");
    }
}

/**
 * Enables the creation of a draft assignment for a project
 */
export class DraftAssignment extends Assignment {

    /**
     * Gets the user who is responsible for entering status for the current assignment
     */
    public get owner(): User {
        return new User(this, "Owner");
    }

    /**
     * Gets the parent assignment link
     */
    public get parent(): DraftAssignment {
        return new DraftAssignment(this, "Parent");
    }

    /**
     * Gets the resource that is associated with the assignment
     */
    public get resource(): DraftProjectResource {
        return new DraftProjectResource(this, "Resource");
    }

    /**
     * Gets the task to which the assignment belongs
     */
    public get task(): DraftTask {
        return new DraftTask(this, "Task");
    }

    /**
     * Deletes the draft assignment
     */
    public delete = this._delete;
}

/**
 * Contains the properties that can be set when creating an assignment
 */
export interface AssignmentCreationInformation {

    /**
     * Gets or sets the assignment finish date and time
     */
    finish?: Date;

    /**
     * Gets or sets the GUID of the assignment
     */
    id?: string;

    /**
     * Gets or sets the notes for the assignment
     */
    notes?: string;

    /**
     * Gets or sets the GUID of the resource for the assignment
     */
    resourceId?: string;

    /**
     * Gets or sets the assignment start date and time
     */
    start?: Date;

    /**
     * Gets or sets the GUID of the task for the assignment
     */
    taskId?: string;
}
