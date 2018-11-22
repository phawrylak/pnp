import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { ProjectDetailPage } from "./projectdetailpages";
import { Stage } from "./stages";

/**
 * Represents a collection of project detail pages (PDPs) that are visible in a workflow stage
 */
export class StageDetailPageCollection extends ProjectQueryableCollection {

    /**
    * Gets a PDP object from the collection of PDPs that are associated with a workflow stage, by using the PDP GUID
    *
    * @param id The string representation of a project detail page GUID
    */
    public getById(id: string): StageDetailPage {
        const page = new StageDetailPage(this);
        page.concat(`('${id}')`);
        return page;
    }

    /**
     * Adds a project detail page to the collection of PDPs that are visible in a workflow stage
     *
     * @param parameters The properties of the stage project detail page to create
     */
    public async add(parameters: StageDetailPageCreationInformation): Promise<CommandResult<StageDetailPage>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a project detail page (PDP) for a workflow stage
 */
export class StageDetailPage extends ProjectQueryableInstance {

    /**
     * Gets the PDP in a workflow stage
     */
    public get page(): ProjectDetailPage {
        return new ProjectDetailPage(this, "Page");
    }

    /**
     * Gets a link to the related workflow stage for the PDP
     */
    public get stage(): Stage {
        return new Stage(this, "Stage");
    }

    /**
     * Deletes the PDP for the workflow stage
     */
    public delete = this._delete;
}

/**
 * Provides property settings and methods that are used to create a project detail page (PDP) for a workflow stage
 */
export interface StageDetailPageCreationInformation {

    /**
     * Gets or sets the description of the stage PDP
     */
    description?: string;

    /**
     * Gets or sets the GUID of the PDP for the stage
     */
    id?: string;

    /**
     * Gets or sets the position of a PDP relative to the other PDPs that are displayed in the stage
     */
    position?: number;

    /**
     * Gets or sets a value that indicates whether the stage PDP requires attention
     */
    requiresAttention?: boolean;
}
