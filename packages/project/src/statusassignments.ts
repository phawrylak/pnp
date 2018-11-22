import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { CustomFieldCollection } from "./customfields";
import { EnterpriseResource } from "./enterpriseresources";
import { PublishedProject } from "./projects";
import { StatusTask, StatusTaskCreationInformation } from "./statustasks";
import { TimePhase } from "./timephases";

/**
 * Represents a collection of StatusAssignment objects, which are assignments in a status update
 */
export class StatusAssignmentCollection extends ProjectQueryableCollection {

    /**
    * Gets a status assignment object from the collection with the specified GUID
    *
    * @param id The string representation of the GUID for the status assignment
    */
    public getById(id: string): StatusAssignment {
        const assignment = new StatusAssignment(this);
        assignment.concat(`('${id}')`);
        return assignment;
    }

    /**
    * Reads the timephased data for assignments that are within the specified start date and end date
    *
    * @param start The start date
    * @param end The end date
    */
    public getTimePhase(start: Date, end: Date): TimePhase {
        const timePhase = new TimePhase(this);
        timePhase.concat(`/GetTimePhaseByUrl('${start.toISOString()}', '${end.toISOString()}')`);
        return timePhase;
    }

    /**
     * Adds the assignment that is specified by the StatusAssignmentCreationInformation object to the collection
     *
     * @param parameters The properties of the status assignment to create
     */
    public async add(parameters: StatusAssignmentCreationInformation): Promise<CommandResult<StatusAssignment>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }

    /**
     * Submit all updates to assignments in the StatusAssignmentCollection for approval
     *
     * @param comment The text comment for the status submission
     */
    public submitAllStatusUpdates(comment: string): Promise<void> {
        return this.clone(StatusAssignmentCollection, "SubmitAllStatusUpdates").postCore({ body: jsS({ comment: comment }) });
    }
}

/**
 * Provides an object that is an assignment in a status update
 */
export class StatusAssignment extends ProjectQueryableInstance {

    /**
     * Gets the collection of custom fields that have values set for the status assignment
     */
    public get customFields(): CustomFieldCollection {
        return new CustomFieldCollection(this, "CustomFields");
    }

    /**
     * Gets the project that contains the status assignment
     */
    public get project(): PublishedProject {
        return new PublishedProject(this, "Project");
    }

    /**
     * Gets the resource that is associated with the status assignment
     */
    public get resource(): EnterpriseResource {
        return new EnterpriseResource(this, "Resource");
    }

    /**
     * Gets the task that is associated with the status assignment
     */
    public get task(): StatusTask {
        return new StatusTask(this, "Task");
    }

    /**
     * Deletes the StatusAssignment object
     */
    public delete = this._delete;

    /**
     * Submits all updates to this assignment for approval
     *
     * @param comment A comment about the submission
     */
    public submitStatusUpdates(comment: string): Promise<void> {
        return this.clone(StatusAssignment, "SubmitStatusUpdates").postCore({ body: jsS({ comment: comment }) });
    }
}

/**
 * Provides property settings and methods for the creation of a status assignment object
 */
export interface StatusAssignmentCreationInformation {

    /**
     * Gets or sets the comment for the status submission
     */
    comment?: string;

    /**
     * Gets or sets the GUID of the status assignment
     */
    id?: string;

    /**
     * Gets or sets the GUID of the project that contains the status assignment
     */
    projectId?: string;

    /**
     * Gets or sets the task parameters for the status assignment
     */
    task?: StatusTaskCreationInformation;
}
