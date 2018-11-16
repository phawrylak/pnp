import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";

/**
 * Represents a collection of PlanAssignmentInterval objects
 */
export class PlanAssignmentIntervalCollection extends ProjectQueryableCollection {

    /**
    * Gets a plan assignment interval from the collection of plan assignment intervals, with the specified object identifier
    *
    * @param id A string value that represents an identifier for a plan assignment interval object
    */
    public getById(id: string): PlanAssignmentInterval {
        const interval = new PlanAssignmentInterval(this);
        interval.concat(`('${id}')`);
        return interval;
    }

    /**
    * Gets a plan assignment interval from the collection of plan assignment intervals, with the specified start date
    *
    * @param start A start date
    */
    public getByStart(start: Date): PlanAssignmentInterval {
        const interval = new PlanAssignmentInterval(this);
        interval.concat(`/GetByStart('${start.toISOString()}')`);
        return interval;
    }
}

/**
 * Represents the collection of time intervals for a project plan assignment
 */
export class PlanAssignmentInterval extends ProjectQueryableInstance {
}

/**
 * Provides information that is used for the creation of PlanAssignmentInterval objects
 */
export interface PlanAssignmentIntervalCreationInformation {

    /**
     * Gets or sets the timespan of the plan assignment interval, expressed in a text string
     */
    duration?: string;

    /**
     * Gets or sets the timescale division that is used to measure the amount of work that is being done
     */
    interval?: Date;
}
