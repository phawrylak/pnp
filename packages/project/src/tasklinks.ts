import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { DraftTask, PublishedTask } from "./tasks";

/**
 * Represents a collection of task links in a published project
 */
export class PublishedTaskLinkCollection extends ProjectQueryableCollection {

    /**
    * Gets a task link from the collection with the specified GUID
    *
    * @param id The string representation of the task link GUID
    */
    public getById(id: string): PublishedTaskLink {
        const taskLink = new PublishedTaskLink(this);
        taskLink.concat(`('${id}')`);
        return taskLink;
    }
}

/**
 * Represents a collection of DraftTaskLink objects
 */
export class DraftTaskLinkCollection extends ProjectQueryableCollection {

    /**
    * Gets a draft task link from the collection with the specified GUID
    *
    * @param id The string representation of the task link GUID
    */
    public getById(id: string): DraftTaskLink {
        const taskLink = new DraftTaskLink(this);
        taskLink.concat(`('${id}')`);
        return taskLink;
    }

    /**
     * Adds the draft task link that is specified by the TaskLinkCreationInformation object to the collection
     *
     * @param parameters The properties of the draft task link to create
     */
    public async add(parameters: TaskLinkCreationInformation): Promise<CommandResult<DraftTaskLink>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents the dependency relationship between the start and finish dates of two tasks
 */
export abstract class TaskLink extends ProjectQueryableInstance {
}

/**
 * Represents a dependency relationship between the start and finish dates of two tasks
 */
export class PublishedTaskLink extends TaskLink {

    /**
     * Gets the task at the end of the link
     */
    public get end(): PublishedTask {
        return new PublishedTask(this, "End");
    }

    /**
     * Gets the task at the start of the link
     */
    public get start(): PublishedTask {
        return new PublishedTask(this, "Start");
    }
}

/**
 * Creates an object to access the task links in a draft project
 */
export class DraftTaskLink extends TaskLink {

    /**
     * Gets the task at the end of the link
     */
    public get end(): DraftTask {
        return new DraftTask(this, "End");
    }

    /**
     * Gets the task at the start of the link
     */
    public get start(): DraftTask {
        return new DraftTask(this, "Start");
    }

    /**
     * Deletes the DraftTaskLink object
     */
    public delete = this._delete;
}

/**
 * Provides property settings and methods that are used to create a task link
 */
export interface TaskLinkCreationInformation {

    /**
     * Gets or sets the type of link relationship between two tasks
     */
    dependencyType?: DependencyType;

    /**
     * Gets or sets the GUID of the task that is at the end of the link
     */
    endId?: string;

    /**
     * Gets or sets the GUID of the task link
     */
    id?: string;

    /**
     * Gets or sets the GUID of the task that is at the start of the link
     */
    startId: string;
}

/**
 * Specifies the type of dependency to establish between two tasks
 */
export enum DependencyType {

    /**
     * Finish to finish (FF) task link type. The successor task B cannot finish until the predecessor task A finishes
     */
    FinishFinish,

    /**
     * This is the default value. Finish to start (FS) task link type. The predecessor task A must finish before the successor task B can start
     */
    FinishStart,

    /**
     * Start to finish (SF) task link type. The predecessor task A must start before the successor task B finishes. This is the least common of the four dependency types
     */
    StartFinish,

    /**
     * Start to start (SS) task link type. The successor task B cannot start until the predecessor task A starts
     */
    StartStart,
}
