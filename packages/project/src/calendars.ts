import { jsS } from "@pnp/common";
import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import { CalendarExceptionCollection } from "./calendarexceptions";

/**
 * Represents a collection of calendars objects
 */
@defaultPath("_api/ProjectServer/Calendars")
export class CalendarCollection extends ProjectQueryableCollection {

    /**
    * Gets a calendar from the collection with the specified GUID
    *
    * @param id The string representation of a calendar GUID
    */
    public getById(id: string): Calendar {
        const calendar = new Calendar(this);
        calendar.concat(`('${id}')`);
        return calendar;
    }

    /**
     * Adds a new calendar to the collection of calendars
     *
     * @param parameters An object that contains information, for example name and GUID, about a new calendar
     */
    public async add(parameters: CalendarCreationInformation): Promise<CommandResult<Calendar>> {
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Represents a Project Server calendar
 */
export class Calendar extends ProjectQueryableInstance {

    /**
     * Gets the collection of exceptions to base calendars
     */
    public get baseCalendarExceptions(): CalendarExceptionCollection {
        return new CalendarExceptionCollection(this, "BaseCalendarExceptions");
    }

    /**
     * Deletes the Calendar object
     */
    public delete = this._delete;

    /**
     * Makes a copy of the calendar and gives the copy a new name
     *
     * @param name The name of the new calendar
     */
    public async copyTo(name: string): Promise<CommandResult<Calendar>> {
        const data = await this.clone(Calendar, "CopyTo").postCore({ body: jsS({ name: name }) });
        return { data: data, instance: this.getParent(CalendarCollection).getById(data.Id) };
    }
}

/**
 * Represents information that is used to create a new calendar
 */
export interface CalendarCreationInformation {

    /**
     * Gets or sets the GUID of the new calendar
     */
    id?: string;

    /**
     * Gets or sets the name of the new calendar
     */
    name?: string;

    /**
     * Gets or sets the original GUID of the new calendar
     */
    originalId?: string;
}
