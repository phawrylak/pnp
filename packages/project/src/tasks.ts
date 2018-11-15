import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";

/**
 * Represents a collection of tasks in a published project
 */
export class PublishedTaskCollection extends ProjectQueryableCollection {

    /**
    * Gets a task in the collection with the specified GUID
    *
    * @param id The string representation of the task GUID
    */
    public getById(id: string): PublishedTask {
        const task = new PublishedTask(this);
        task.concat(`('${id}')`);
        return task;
    }
}

/**
 * Represents a collection of DraftTask objects
 */
export class DraftTaskCollection extends ProjectQueryableCollection {

    /**
    * Gets a draft task from the collection with the specified GUID
    *
    * @param id The string representation of the task GUID
    */
    public getById(id: string): DraftTask {
        const task = new DraftTask(this);
        task.concat(`('${id}')`);
        return task;
    }

    /**
     * Adds the draft task that is specified by the TaskCreationInformation object to the collection
     *
     * @param parameters The properties of the draft task to create
     */
    public async add(parameters: TaskCreationInformation): Promise<DraftTaskResult> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, draft: this.getById(data.Id) };
    }
}

/**
 * Contains methods and properties that can be used to access the details of the task
 */
export abstract class Task extends ProjectQueryableInstance {
}

/**
 * Represents a task in a published project
 */
export class PublishedTask extends Task {
}

/**
 * Represents a task in a checked-out project
 */
export class DraftTask extends Task {

    /**
     * Deletes the DraftTask object
     */
    public delete = this._delete;
}

/**
 * Provides property settings and methods that are used to create a task
 */
export interface TaskCreationInformation {

    /**
     * Gets or sets the GUID value of the task that specifies the insertion point
     */
    addAfterId?: string;

    /**
     * Gets or sets the task duration
     */
    duration?: string;

    /**
     * Gets or sets the task finish date
     */
    finish?: Date;

    /**
     * Gets or sets the GUID of the task
     */
    id?: string;

    /**
     * Gets or sets a value that indicates whether the task is manually scheduled
     */
    isManual?: boolean;

    /**
     * Gets or sets the task name
     */
    name: string;

    /**
     * Gets or sets notes about the task
     */
    notes?: string;

    /**
     * Gets or sets the GUID of the parent task in a hierarchical task list
     */
    parentId?: string;

    /**
     * Gets or sets the task start date
     */
    start?: Date;
}

export interface DraftTaskResult {
    data: any;
    draft: DraftTask;
}
