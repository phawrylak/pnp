import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";

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
    public async add(parameters: AssignmentCreationInformation): Promise<DraftAssignmentResult> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, draft: this.getById(data.Id) };
    }
}

/**
 * Contains the common properties for draft assignments and published assignments
 */
export abstract class Assignment extends ProjectQueryableInstance {
}

/**
 * Represents the assignment that is in a published project
 */
export class PublishedAssignment extends Assignment {
}

/**
 * Enables the creation of a draft assignment for a project
 */
export class DraftAssignment extends Assignment {

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

export interface DraftAssignmentResult {
    data: any;
    draft: DraftAssignment;
}
