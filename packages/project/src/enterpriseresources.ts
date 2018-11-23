import { jsS, TypedHash } from "@pnp/common";
import { wrap } from "./utils/wrap";
import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { CalendarExceptionCollection } from "./calendarexceptions";
import { Calendar } from "./calendars";
import { CustomFieldCollection } from "./customfields";
import { StatusAssignmentCollection } from "./statusassignments";
import { User } from "./users";

/**
 * Represents a collection of EnterpriseResource objects
 */
@defaultPath("_api/ProjectServer/EnterpriseResources")
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
        const data = await this.clone(EnterpriseResourceCollection, "Add").postCore({ body: jsS(wrap({parameters: parameters})) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a resource that is managed by Project Server in a project
 */
export class EnterpriseResource extends ProjectQueryableInstance {

    /**
     * Gets a collection of status assignments for an enterprise resource
     */
    public get assignments(): StatusAssignmentCollection {
        return new StatusAssignmentCollection(this, "Assignments");
    }

    /**
     * Gets a base calendar for an enterprise resource
     */
    public get baseCalendar(): Calendar {
        return new Calendar(this, "BaseCalendar");
    }

    /**
     * Gets a collection of custom fields that have values set for an enterprise resource
     */
    public get customFields(): CustomFieldCollection {
        return new CustomFieldCollection(this, "CustomFields");
    }

    /**
     * Gets the default user that is entered into the Assignment Owner field when an enterprise resource is assigned to a task
     */
    public get defaultAssignmentOwner(): User {
        return new User(this, "DefaultAssignmentOwner");
    }

    /**
     * Gets a collection of exceptions to the base calendar that are specific to an enterprise resource
     */
    public get resourceCalendarExceptions(): CalendarExceptionCollection {
        return new CalendarExceptionCollection(this, "ResourceCalendarExceptions");
    }

    /**
     * Gets the manager who reviews and approves the timesheet of an enterprise resource
     */
    public get timesheetManager(): User {
        return new User(this, "TimesheetManager");
    }

    /**
     * Gets the SharePoint user that is linked to the Enterprise Resource
     */
    public get user(): User {
        return new User(this, "User");
    }

    /**
     * Deletes the EnterpriseResource object
     */
    public delete = this._delete;

    /**
    * Updates this enterprise resource with the supplied properties
    *
    * @param properties A plain object of property names and values to update the enterprise resource
    */
    public update = this._update<void, TypedHash<any>, any>(
        "PS.EnterpriseResource",
        _ => null);

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
    HyperlinkName?: string;

    /**
     * Gets or sets the hyperlink URL of the enterprise resource
     */
    HyperlinkUrl?: string;

    /**
     * Gets or sets the GUID of the enterprise resource
     */
    Id?: string;

    /**
     * Gets or sets a Boolean value that indicates whether this is a budget resource
     */
    IsBudget?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether this is a generic resource
     */
    IsGeneric?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether this resource should be created in an inactive state
     */
    IsInactive?: boolean;

    /**
     * Gets or sets the name of the enterprise resource
     */
    Name: string;

    /**
     * Gets or sets a value that represents the resource type
     */
    ResourceType?: EnterpriseResourceType;
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
