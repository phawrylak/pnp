import { jsS } from "@pnp/common";
import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { EntityType } from "./entitytypes";
import { LookupEntryCollection } from "./lookupentries";
import { LookupTable } from "./lookuptables";

/**
 * Represents a collection of CustomField objects
 */
@defaultPath("_api/ProjectServer/CustomFields")
export class CustomFieldCollection extends ProjectQueryableCollection {

    /**
    * Gets a custom field from the collection with the specified GUID
    *
    * @param id The string representation of the custom field GUID
    */
    public getById(id: string): CustomField {
        const customField = new CustomField(this);
        customField.concat(`('${id}')`);
        return customField;
    }

    /**
    * Gets a custom field from the collection by using the alternate custom field GUID that is specified in an App package for Project Online
    *
    * @param objectId The alternate custom field GUID
    */
    public getByAppAlternateId(objectId: string): CustomField {
        const customField = new CustomField(this);
        customField.concat(`/GetByAppAlternateId('${objectId}')`);
        return customField;
    }

    /**
     * Adds the custom field that is specified by the CustomFieldCreationInformation object to the collection
     *
     * @param parameters An object that contains the information for the creation of a custom field
     */
    public async add(parameters: CustomFieldCreationInformation): Promise<CommandResult<CustomField>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Contains the properties and methods that are used to create an enterprise custom field
 */
export class CustomField extends ProjectQueryableInstance {

    /**
     * Gets the type of entity (project, task, or resource) for the custom field
     */
    public get entityType(): EntityType {
        return new EntityType(this, "EntityType");
    }

    /**
     * Gets a collection of valid lookup table entries for this field
     */
    public get lookupEntries(): LookupEntryCollection {
        return new LookupEntryCollection(this, "LookupEntries");
    }

    /**
     * Gets the LookupTable object for a text custom field
     */
    public get lookupTable(): LookupTable {
        return new LookupTable(this, "LookupTable");
    }

    /**
     * Deletes the CustomField object
     */
    public delete = this._delete;
}

/**
 * Provides information that is used in the creation of a custom field
 */
export interface CustomFieldCreationInformation {

    /**
     * Gets or sets the description of the custom field
     */
    description?: string;

    /**
     * Gets or sets the type of entity (project, task, or resource) for the custom field
     */
    entityTypeId?: string;

    /**
     * Gets or sets the type of the custom field
     */
    fieldType?: CustomFieldType;

    /**
     * Gets or sets the formula that calculates the value of a custom field
     */
    formula?: string;

    /**
     * Gets or sets the graphical indicator non summary
     */
    graphicalIndicatorNonSummary?: string;

    /**
     * Gets or sets the graphical indicator project summary
     */
    graphicalIndicatorProjectSummary?: string;

    /**
     * Gets or sets the graphical indicator summary
     */
    graphicalIndicatorSummary?: string;

    /**
     * Gets or sets the GUID of the custom field
     */
    id?: string;

    /**
     * Gets or sets a Boolean value that indicates whether the custom field value can be edited in project site Visibility mode
     */
    isEditableInVisibility?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether the custom field is leaf only
     */
    isLeafOnly?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether a text custom field can contain multiple lines
     */
    isMultilineText?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether a value for the custom field is required when the entity (project, resource, or task) is created
     */
    isRequired?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether the custom field value is controlled by a workflow
     */
    isWorkflowControlled?: boolean;

    /**
     * Gets or sets a Boolean value that indicate whether multiple lookup table values can be set for this custom field
     */
    lookupAllowMultiSelect?: boolean;

    /**
     * Gets or sets the GUID of the default lookup table value
     */
    lookupDefaultValue?: string;

    /**
     * Gets or sets the LookupTable GUID for a text custom field
     */
    lookupTableId?: string;

    /**
     * Gets or sets the name of the custom field
     */
    name: string;

    /**
     * Gets or sets a Boolean value that indicates whether tooltip should be displayed
     */
    showToolTip?: boolean;

    /**
     * Gets or sets a Boolean value that indicates whether field should be used for matching
     */
    useFieldForMatching?: boolean;
}

/**
 * Specifies the type for an enterprise custom field
 */
export enum CustomFieldType {

    /**
     * A date value. HIWORD contains days offset from 1/1/84. LOWORD contains minute off-set, ranging from 0 to 1440, from 12:00 A.M. (midnight).
     */
    DATE = 4,

    /**
     * Value in 1/10 minutes.
     */
    DURATION = 6,

    /**
     * Value in 1/100 dollars.
     */
    COST = 9,

    /**
     * A number value.
     */
    NUMBER = 15, // 0x0000000F

    /**
     * Index into yes/no string table.
     */
    FLAG = 17, // 0x00000011

    /**
     * A string value.
     */
    TEXT = 21, // 0x00000015

    /**
     * A date value; if no time is included, the default finish time is used. HIWORD contains days offset from 1/1/84.
     * LOWORD contains minute off-set, ranging from 0 to 1440, from 12:00 A.M. (midnight).
     */
    FINISHDATE = 27, // 0x0000001B
}
