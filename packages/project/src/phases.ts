import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";

/**
 * Represents a collection of workflow Phase objects
 */
export class PhaseCollection extends ProjectQueryableCollection {

    /**
    * Gets a workflow phase from the collection with the specified GUID
    *
    * @param id The string representation of the workflow phase GUID
    */
    public getById(id: string): Phase {
        const phase = new Phase(this);
        phase.concat(`('${id}')`);
        return phase;
    }

    /**
     * Adds the workflow phase that is specified by the PhaseCreationInformation object to the collection
     *
     * @param parameters The properties of the workflow phase to create
     */
    public async add(parameters: PhaseCreationInformation): Promise<CommandResult<Phase>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a collection of stages that are grouped to identify a common set of activities in the project life cycle
 */
export class Phase extends ProjectQueryableInstance {

    /**
     * Deletes the Phase object
     */
    public delete = this._delete;
}

/**
 * Provides methods and property settings that are used in the creation of a workflow phase
 */
export interface PhaseCreationInformation {

    /**
     * Gets or sets the description of the phase
     */
    description?: string;

    /**
     * Gets or sets the GUID for the phase
     */
    id?: string;

    /**
     * Gets or sets the name of the phase
     */
    name: string;
}
