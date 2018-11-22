import { ProjectQueryableInstance } from "./projectqueryable";
import { StatusAssignmentCollection } from "./statusassignments";

/**
 * Represents assignment progress information that is distributed over time
 */
export class TimePhase extends ProjectQueryableInstance {

    /**
     * TODO
     */
    public get assignments(): StatusAssignmentCollection {
        return new StatusAssignmentCollection(this, "Assignments");
    }
}
