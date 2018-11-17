import { jsS, TypedHash } from "@pnp/common";
import { ProjectQueryableInstance } from "./projectqueryable";

/**
 * Contains the methods and properties for managing a timesheet
 */
export class TimeSheet extends ProjectQueryableInstance {

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
    public submit(comment?: string): Promise<void> {
        return this.clone(TimeSheet, "Submit").postCore({ body: jsS({ comment: comment }) });
    }
}
