import { jsS, TypedHash } from "@pnp/common";
import { wrap } from "./utils/wrap";
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
        const data = await this.clone(LookupTableCollection, "Add").postCore({ body: jsS(wrap({parameters: parameters})) });
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

    /**
    * Updates this lookup table with the supplied properties
    *
    * @param properties A plain object of property names and values to update the lookup table
    */
    public update = this._update<void, TypedHash<any>, any>(
        "PS.LookupTable",
        _ => null);
}

/**
 * Provides methods and property settings for the creation of a lookup table
 */
export interface LookupTableCreationInformation {

    /**
     * Gets or sets the collection of entries in the lookup table
     */
    Entries: LookupEntryCreationInformation[];

    /**
     * Gets or sets the GUID of the lookup table
     */
    Id?: string;

    /**
     * Gets or sets the collection of mask definitions for the levels of a hierarchical lookup table
     */
    Masks: LookupMask[];

    /**
     * Gets or sets the name of the lookup table
     */
    Name: string;

    /**
     * Gets or sets the sort order for the entries in the table
     */
    SortOrder?: LookupTableSortOrder;
}

/**
 * Represents a mask definition for the levels of a hierarchical lookup table
 */
export interface LookupMask {

    /**
     * Gets or sets the number of characters in the mask sequence
     */
    Length?: number;

    /**
     * Gets or sets the mask segment type
     */
    MaskType?: LookupTableMaskSequence;

    /**
     * Gets or sets the separator string that is used in the concatenation of parent table entry values at each level of the table mask
     */
    Separator: string;
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
