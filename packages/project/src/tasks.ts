import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { DraftAssignmentCollection, PublishedAssignmentCollection } from "./assignments";
import { Calendar } from "./calendars";
import { CustomFieldCollection } from "./customfields";
import { PublishedProject } from "./projects";
import { DraftTaskLinkCollection, PublishedTaskLinkCollection } from "./tasklinks";
import { User } from "./users";

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
    public async add(parameters: TaskCreationInformation): Promise<CommandResult<DraftTask>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Contains methods and properties that can be used to access the details of the task
 */
export abstract class Task extends ProjectQueryableInstance {

    /**
     * TODO
     */
    public get customFields(): CustomFieldCollection {
        return new CustomFieldCollection(this, "CustomFields");
    }

    /**
     * TODO
     */
    public get subProject(): PublishedProject {
        return new PublishedProject(this, "SubProject");
    }
}

/**
 * Represents a task in a published project
 */
export class PublishedTask extends Task {

    /**
     * TODO
     */
    public get assignments(): PublishedAssignmentCollection {
        return new PublishedAssignmentCollection(this, "Assignments");
    }

    /**
     * TODO
     */
    public get calendar(): Calendar {
        return new Calendar(this, "Calendar");
    }

    /**
     * TODO
     */
    public get parent(): PublishedTask {
        return new PublishedTask(this, "Parent");
    }

    /**
     * TODO
     */
    public get predecessors(): PublishedTaskLinkCollection {
        return new PublishedTaskLinkCollection(this, "Predecessors");
    }

    /**
     * TODO
     */
    public get statusManager(): User {
        return new User(this, "StatusManager");
    }

    /**
     * TODO
     */
    public get successors(): PublishedTaskLinkCollection {
        return new PublishedTaskLinkCollection(this, "Successors");
    }
}

/**
 * Represents a task in a checked-out project
 */
export class DraftTask extends Task {

    /**
     * TODO
     */
    public get assignments(): DraftAssignmentCollection {
        return new DraftAssignmentCollection(this, "Assignments");
    }

    /**
     * TODO
     */
    public get calendar(): Calendar {
        return new Calendar(this, "Calendar");
    }

    /**
     * TODO
     */
    public get parent(): DraftTask {
        return new DraftTask(this, "Parent");
    }

    /**
     * TODO
     */
    public get predecessors(): DraftTaskLinkCollection {
        return new DraftTaskLinkCollection(this, "Predecessors");
    }

    /**
     * TODO
     */
    public get statusManager(): User {
        return new User(this, "StatusManager");
    }

    /**
     * TODO
     */
    public get successors(): DraftTaskLinkCollection {
        return new DraftTaskLinkCollection(this, "Successors");
    }

    /**
     * Deletes the DraftTask object
     */
    public delete = this._delete;
}

/**
 * TODO
 */
export class ProjectSummaryTask extends Task {
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
