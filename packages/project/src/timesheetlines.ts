import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { PublishedAssignment } from "./assignments";
import { TimeSheet } from "./timesheets";
import { TimeSheetWorkCollection } from "./timesheetworks";

/**
 * Represents a collection of timesheet lines
 */
export class TimeSheetLineCollection extends ProjectQueryableCollection {

    /**
    * Gets the timesheet line from the collection with the specified GUID
    *
    * @param id The string representation of the timesheet line GUID
    */
    public getById(id: string): TimeSheetLine {
        const line = new TimeSheetLine(this);
        line.concat(`('${id}')`);
        return line;
    }

    /**
     * Adds the specified timesheet line to the timesheet line collection
     *
     * @param parameters The properties of the timesheet line to create
     */
    public async add(parameters: TimeSheetLineCreationInformation): Promise<CommandResult<TimeSheetLine>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a line in a timesheet
 */
export class TimeSheetLine extends ProjectQueryableInstance {

    /**
     * Gets the assignment that is associated with the timesheet line
     */
    public get assignment(): PublishedAssignment {
        return new PublishedAssignment(this, "Assignment");
    }

    /**
     * Gets the timesheet that is associated with the timesheet line
     */
    public get timeSheet(): TimeSheet {
        return new TimeSheet(this, "TimeSheet");
    }

    /**
     * Gets the actual work on the timesheet
     */
    public get work(): TimeSheetWorkCollection {
        return new TimeSheetWorkCollection(this, "Work");
    }

    /**
     * Deletes the TimeSheetLine object
     */
    public delete = this._delete;

    /**
     * Submits the time sheet line
     *
     * @param comment A comment about the timesheet line
     */
    public submit(comment: string): Promise<void> {
        return this.clone(TimeSheetLine, "Submit").postCore({ body: jsS({ comment: comment }) });
    }
}

/**
 * Provides property settings and methods that are used to create a timesheet line
 */
export interface TimeSheetLineCreationInformation {

    /**
     * Gets or sets the GUID of the assignment that is associated with the timesheet line
     */
    assignmentId?: string;

    /**
     * Gets or sets the comment for the timesheet line
     */
    comment?: string;

    /**
     * Gets or sets the GUID for the timesheet line
     */
    id?: string;

    /**
     * Gets or sets the line class type of the timesheet line
     */
    lineClass?: TimeSheetLineClass;

    /**
     * Gets or sets the GUID of the project that is associated with the timesheet line
     */
    projectId?: string;

    /**
     * Gets or sets the time sheet line task name
     */
    taskName?: string;
}

/**
 * Represents classifications that define the different uses of a timesheet line
 */
export enum TimeSheetLineClass {

    /**
     * The class of timesheet line that is used to record standard work time; the default class of timesheet line
     */
    StandardLine,

    /**
     * The class of timesheet line that is used to record sick time
     */
    SickTimeLine,

    /**
     * The class of timesheet line that is used to record vacation time
     */
    VacationLine,

    /**
     * The class of timesheet line that is used to record administrative time
     */
    AdministrativeLine,
}
