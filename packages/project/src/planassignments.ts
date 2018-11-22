import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { CustomFieldCollection } from "./customfields";
import { EnterpriseResource } from "./enterpriseresources";
import {
    PlanAssignmentIntervalCollection,
    PlanAssignmentIntervalCreationInformation,
} from "./planassignmentintervals";

/**
 * Represents a collection of plan assignment objects
 */
export class PlanAssignmentCollection extends ProjectQueryableCollection {

    /**
    * Gets a plan assignment from the collection, by using the specified GUID
    *
    * @param id The string representation of the plan assignment GUID
    */
    public getById(id: string): PlanAssignment {
        const assignment = new PlanAssignment(this);
        assignment.concat(`('${id}')`);
        return assignment;
    }

    /**
     * Adds a new plan assignment to the collection of plan assignments
     *
     * @param parameters An object that contains information, for example name and GUID, about a new plan assignment
     */
    public async add(parameters: PlanAssignmentCreationInformation): Promise<CommandResult<PlanAssignment>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Provides information about an assignment in a project plan
 */
export class PlanAssignment extends ProjectQueryableInstance {

    /**
     * Gets the collection of custom fields that have values set for a plan assignment
     */
    public get customFields(): CustomFieldCollection {
        return new CustomFieldCollection(this, "CustomFields");
    }

    /**
     * Gets the collection of plan assignment time intervals
     */
    public get intervals(): PlanAssignmentIntervalCollection {
        return new PlanAssignmentIntervalCollection(this, "Intervals");
    }

    /**
     * Gets a resource that is part of an organization's entire list of resources and that can be shared across projects
     */
    public get resource(): EnterpriseResource {
        return new EnterpriseResource(this, "Resource");
    }

    /**
     * Deletes the PlanAssignment object
     */
    public delete = this._delete;
}

/**
 * Provides information for the creation of a PlanAssignment object
 */
export interface PlanAssignmentCreationInformation {

    /**
     * Gets or sets the booking type of the assignment
     */
    bookingType?: BookingType;

    /**
     * Gets or sets the GUID for the plan assignment
     */
    id?: string;

    /**
     * Gets or sets an enumerator that iterates through a collection of time intervals
     */
    intervals?: PlanAssignmentIntervalCreationInformation[];

    /**
     * Gets or sets the GUID of the resource
     */
    resourceId?: string;
}

/**
 * Specifies how resources are booked for assignments
 */
export enum BookingType {

    /**
     * The resource booking type is not specified. This is the default value
     */
    NotSpecified,

    /**
     * Resources are booked as "Committed."
     */
    Committed,

    /**
     * Resources are booked as "Proposed."
     */
    Proposed,
}
