import { jsS, TypedHash } from "@pnp/common";
import { ProjectQueryableInstance } from "./projectqueryable";
import { TimeSheetLineCollection } from "./timesheetlines";
import { TimeSheetPeriod } from "./timesheetperiods";
import { User } from "./users";

/**
 * Contains the methods and properties for managing a timesheet
 */
export class TimeSheet extends ProjectQueryableInstance {

    /**
     * Gets the timesheet creator
     */
    public get creator(): User {
        return new User(this, "Creator");
    }

    /**
     * Gets the collection of lines in the timesheet
     */
    public get lines(): TimeSheetLineCollection {
        return new TimeSheetLineCollection(this, "Lines");
    }

    /**
     * Gets the timesheet manager
     */
    public get manager(): User {
        return new User(this, "Manager");
    }

    /**
     * Gets the timesheet period
     */
    public get period(): TimeSheetPeriod {
        return new TimeSheetPeriod(this, "Period");
    }

    /**
    * Updates this timesheet with the supplied properties
    *
    * @param properties A plain object of property names and values to update the timesheet
    */
    public update = this._update<void, TypedHash<any>, any>(
        "PS.TimeSheet",
        _ => null);

    /**
     * Deletes the TimeSheet object
     */
    public delete = this._delete;

    /**
     * Recalls the timesheet
     */
    public recall(): Promise<void> {
        return this.clone(TimeSheet, "Recall").postCore();
    }

    /**
     * Submits the timesheet
     *
     * @param comment A comment on the timesheet
     */
    public submit(comment: string): Promise<void> {
        return this.clone(TimeSheet, "Submit").postCore({ body: jsS({ comment: comment }) });
    }
}
