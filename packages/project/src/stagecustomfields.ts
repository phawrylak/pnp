import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { Stage } from "./stages";

/**
 * Represents a collection of StageCustomField objects, which are custom fields in a workflow stage
 */
export class StageCustomFieldCollection extends ProjectQueryableCollection {

    /**
    * Gets a stage custom field from the collection with the specified GUID
    *
    * @param id The string representation of the stage custom field GUID
    */
    public getById(id: string): StageCustomField {
        const field = new StageCustomField(this);
        field.concat(`('${id}')`);
        return field;
    }

    /**
     * Adds the custom field that is specified by the StageCustomFieldCreationInformation object to the collection
     *
     * @param parameters The properties of the custom field to add to the stage
     */
    public async add(parameters: StageCustomFieldCreationInformation): Promise<CommandResult<StageCustomField>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a custom field for a project stage
 */
export class StageCustomField extends ProjectQueryableInstance {

    /**
     * Gets a link to the Stage entity
     */
    public get stage(): Stage {
        return new Stage(this, "Stage");
    }

    /**
     * Deletes the StageCustomField object
     */
    public delete = this._delete;
}

/**
 * Provides property settings and methods that are used to add a custom field to project stage information
 */
export interface StageCustomFieldCreationInformation {

    /**
     * Gets or sets the GUID for the custom field
     */
    id?: string;

    /**
     * Gets or sets a value that indicates whether the custom field is read-only in the stage
     */
    readOnly?: boolean;

    /**
     * Gets or sets a value that indicates whether the custom field is required in the stage
     */
    required?: boolean;
}
