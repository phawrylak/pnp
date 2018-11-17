import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";

/**
 * Represents a collection of LookupEntry objects for a lookup table
 */
export class LookupEntryCollection extends ProjectQueryableCollection {

    /**
    * Gets a lookup entry from the collection with the specified GUID
    *
    * @param id The string representation of the lookup entry GUID
    */
    public getById(id: string): LookupEntry {
        const entry = new LookupEntry(this);
        entry.concat(`('${id}')`);
        return entry;
    }

    /**
    * Gets a lookup entry from the collection by using the alternate lookup entry GUID that is specified in an App package for Project Server online
    *
    * @param objectId A string object identifier
    */
    public getByAppAlternateId(objectId: string): LookupEntry {
        const entry = new LookupEntry(this);
        entry.concat(`/GetByAppAlternateId('${objectId}')`);
        return entry;
    }

    /**
     * Adds the lookup entry that is specified by the LookupEntryCreationInformation object to the collection
     *
     * @param parameters The properties of the lookup entry to create
     */
    public async add(parameters: LookupEntryCreationInformation): Promise<CommandResult<LookupEntry>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a lookup table entry
 */
export class LookupEntry extends ProjectQueryableInstance {

    /**
     * Deletes the LookupEntry object
     */
    public delete = this._delete;
}

/**
 * Provides information for the creation of a lookup table entry
 */
export interface LookupEntryCreationInformation {

    /**
     * Gets or sets the description of the lookup table entry
     */
    description?: string;

    /**
     * Gets or sets the GUID of the lookup table entry
     */
    id?: string;

    /**
     * Gets or sets the GUID of the parent lookup table entry in a hierarchical text lookup table
     */
    parentId?: string;

    /**
     * Gets or sets an index number for the lookup table entry
     */
    sortIndex?: number;

    /**
     * Gets or sets the value of the lookup table entry
     */
    value?: LookupEntryValue;
}

/**
 * Represents the value of a lookup table entry
 */
export interface LookupEntryValue {

    /**
     * Gets or sets a date and time entry value
     */
    dateValue?: Date;

    /**
     * Gets or sets a duration entry value as a string
     */
    durationValue?: string;

    /**
     * Gets or sets a decimal number entry value
     */
    numberValue?: number;

    /**
     * Gets or sets a text entry value
     */
    textValue?: string;
}
