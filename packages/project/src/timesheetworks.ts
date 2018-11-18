import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";

/**
 * Provides a collection of actual work entries for a timesheet
 */
export class TimeSheetWorkCollection extends ProjectQueryableCollection {

    /**
    * Gets a timesheet actual work object from the collection with the specified GUID
    *
    * @param id The string representation of a timesheet actual work GUID
    */
    public getById(id: string): TimeSheetWork {
        const work = new TimeSheetWork(this);
        work.concat(`('${id}')`);
        return work;
    }

    /**
    * Gets a timesheet actual work object from the collection with a specified start date
    *
    * @param start The start date and time
    */
    public getByStartDate(start: Date): TimeSheetWork {
        const work = new TimeSheetWork(this);
        work.concat(`/GetByStartDate('${start.toISOString()}')`);
        return work;
    }

    /**
     * Adds an actual work value to a timesheet
     *
     * @param parameters An object that contains information for the creation of an actual work item
     */
    public async add(parameters: TimeSheetWorkCreationInformation): Promise<CommandResult<TimeSheetWork>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents the different types of work on a timesheet
 */
export class TimeSheetWork extends ProjectQueryableInstance {

    /**
     * Deletes the TimeSheetWork object
     */
    public delete = this._delete;
}

/**
 * Provides property settings and methods that are used to create a timesheet work object
 */
export interface TimeSheetWorkCreationInformation {

    /**
     * Gets or sets the amount of actual work that is on a timesheet
     */
    actualWork?: string;

    /**
     * Gets or sets the timesheet work comment
     */
    comment?: string;

    /**
     * Gets or sets the end time of work that is on a timesheet
     */
    end?: Date;

    /**
     * Gets or sets the amount of non-billable overtime work that is on a timesheet
     */
    nonBillableOvertimeWork?: string;

    /**
     * Gets or sets the amount of non-billable work that is on a timesheet
     */
    nonBillableWork?: string;

    /**
     * Gets or sets the amount of overtime work that is on a timesheet
     */
    overtimeWork?: string;

    /**
     * Gets or sets the amount of planned work that is on a timesheet
     */
    plannedWork?: string;

    /**
     * Gets or sets the start time of work that is on a timesheet
     */
    start?: Date;
}
