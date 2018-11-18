import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { TimeSheet } from "./timesheets";

/**
 * Represents a collection of TimeSheetPeriod objects
 */
export class TimeSheetPeriodCollection extends ProjectQueryableCollection {

    /**
    * Gets a timesheet period from the collection with the specified GUID
    *
    * @param id The string representation of the timesheet period GUID
    */
    public getById(id: string): TimeSheetPeriod {
        const period = new TimeSheetPeriod(this);
        period.concat(`('${id}')`);
        return period;
    }
}

/**
 * Represents a defined period of time on a timesheet
 */
export class TimeSheetPeriod extends ProjectQueryableInstance {

    /**
     * Creates a new TimeSheet object
     */
    public createTimeSheet(): Promise<CommandResult<TimeSheet>> {
        return this.clone(TimeSheetPeriod, "CreateTimeSheet").postCore();
    }
}
