import { jsS } from "@pnp/common";
import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import {
    LookupEntryCollection,
    LookupEntryCreationInformation,
} from "./lookupentries";

/**
 * Represents a collection of LookupTable objects
 */
@defaultPath("_api/ProjectServer/LookupTables")
export class LookupTableCollection extends ProjectQueryableCollection {

    /**
    * Gets a lookup table from the collection with the specified GUID
    *
    * @param id The string representation of the lookup table GUID
    */
    public getById(id: string): LookupTable {
        const table = new LookupTable(this);
        table.concat(`('${id}')`);
        return table;
    }

    /**
    * Get an element from the lookup table collection by using the alternate object GUID that is specified in an App package for Project Online
    *
    * @param objectId The alternate custom field GUID
    */
    public getByAppAlternateId(objectId: string): LookupTable {
        const table = new LookupTable(this);
        table.concat(`/GetByAppAlternateId('${objectId}')`);
        return table;
    }

    /**
     * Adds the lookup table that is specified by the LookupTableCreationInformation object to the collection
     *
     * @param parameters The properties of the lookup table to create
     */
    public async add(parameters: LookupTableCreationInformation): Promise<CommandResult<LookupTable>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a lookup table
 */
export class LookupTable extends ProjectQueryableInstance {

    /**
     * Gets the collection of entries in the lookup table
     */
    public get entries(): LookupEntryCollection {
        return new LookupEntryCollection(this, "Entries");
    }

    /**
     * Deletes the LookupTable object
     */
    public delete = this._delete;
}

/**
 * Provides methods and property settings for the creation of a lookup table
 */
export interface LookupTableCreationInformation {

    /**
     * Gets or sets the collection of entries in the lookup table
     */
    entries?: LookupEntryCreationInformation[];

    /**
     * Gets or sets the GUID of the lookup table
     */
    id?: string;

    /**
     * Gets or sets the collection of mask definitions for the levels of a hierarchical lookup table
     */
    masks?: LookupMask[];

    /**
     * Gets or sets the name of the lookup table
     */
    name: string;

    /**
     * Gets or sets the sort order for the entries in the table
     */
    sortOrder?: LookupTableSortOrder;
}

/**
 * Represents a mask definition for the levels of a hierarchical lookup table
 */
export interface LookupMask {

    /**
     * Gets or sets the number of characters in the mask sequence
     */
    length?: number;

    /**
     * Gets or sets the mask segment type
     */
    maskType?: LookupTableMaskSequence;

    /**
     * Gets or sets the separator string that is used in the concatenation of parent table entry values at each level of the table mask
     */
    separator?: string;
}

/**
 * Specifies the mask sequence, which is the type of data for a lookup table
 */
export enum LookupTableMaskSequence {

    /**
     * Number data expressed as a string
     */
    NUMBER_TEXT,

    /**
     * Uppercase character data
     */
    UPPERCASE,

    /**
     * Lowercase character data
     */
    LOWERCASE,

    /**
     * Mixed uppercase and lowercase character data
     */
    CHARACTERS,

    /**
     * Date data
     */
    DATE,

    /**
     * Cost data
     */
    COST,

    /**
     * Duration data
     */
    DURATION,

    /**
     * Decimal number data
     */
    NUMBER_DECIMAL,

    /**
     * Flag data (yes/no values)
     */
    FLAG,
}

/**
 * Specifies the sort order for a lookup table
 */
export enum LookupTableSortOrder {

    /**
     * Custom sort; the user specifies the preferred ordering of lookup table values
     */
    UserDefined,

    /**
     * Ascending sort; for example, A comes before B
     */
    Ascending,

    /**
     * Descending sort; for example, B comes before A
     */
    Descending,
}
