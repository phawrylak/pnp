import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";

/**
 * Represents a collection of EnterpriseResource objects
 */
export class EnterpriseResourceCollection extends ProjectQueryableCollection {

    /**
    * Gets an enterprise resource from the collection with the specified GUID
    *
    * @param id The string representation of the enterprise resource GUID
    */
    public getById(id: string): EnterpriseResource {
        const resource = new EnterpriseResource(this);
        resource.concat(`('${id}')`);
        return resource;
    }

    /**
     * Adds the enterprise resource that is specified by the AssignmentCreationInformation object to the collection
     *
     * @param parameters The properties of the enterprise resource to create
     */
    public async add(parameters: EnterpriseResourceCreationInformation): Promise<CommandResult<EnterpriseResource>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a resource that is managed by Project Server in a project
 */
export class EnterpriseResource extends ProjectQueryableInstance {

    /**
     * Deletes the EnterpriseResource object
     */
    public delete = this._delete;

    /**
     * Forces a project to be checked in after it is left in a state of being checked out following the interruption or unexpected closing of Project Server
     */
    public forceCheckIn(): Promise<void> {
        return this.clone(EnterpriseResource, "ForceCheckIn").postCore();
    }
}

/**
 * Provides information for the creation of an enterprise resource
 */
export interface EnterpriseResourceCreationInformation {

    /**
     * Gets or sets the hyperlink name of the enterprise resource
     */
    hyperlinkName?: string;

    /**
     * Gets or sets the hyperlink URL of the enterprise resource
     */
    hyperlinkUrl?: string;

    /**
     * Gets or sets the GUID of the enterprise resource
     */
    id?: string;

    /**
     * Gets or sets a Boolean value that indicates whether this is a budget resource
     */
    isBudget?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether this is a generic resource
     */
    isGeneric?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether this resource should be created in an inactive state
     */
    isInactive?: boolean;

    /**
     * Gets or sets the name of the enterprise resource
     */
    name: string;

    /**
     * Gets or sets a value that represents the resource type
     */
    resourceType?: EnterpriseResourceType;
}

/**
 * Represents the different types of enterprise resources
 */
export enum EnterpriseResourceType {

    /**
     * Type is not specified
     */
    NotSpecified,

    /**
     * People and equipment resources that perform work to accomplish a task
     */
    Work,

    /**
     * The supplies or other consumable items that are used to complete tasks in a project
     */
    Material,

    /**
     * The cost that must be incurred to complete tasks in a project
     */
    Cost,
}
