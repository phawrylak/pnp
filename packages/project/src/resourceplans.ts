import { TypedHash } from "@pnp/common";
import { ProjectQueryableInstance } from "./projectqueryable";
import { DraftProject } from "./projects";
import { QueueJob } from "./queuejobs";
import { CommandResult } from "./types";
import { PlanAssignmentCollection } from "./planassignments";

/**
 * Represents a high-level look at what resources might be needed for a project
 */
export class ResourcePlan extends ProjectQueryableInstance {

    /**
     * Gets a collection of PlanAssignment objects that are associated with the resource plan
     */
    public get assignments(): PlanAssignmentCollection {
        return new PlanAssignmentCollection(this, "Assignments");
    }

    /**
    * Updates this resource plan with the supplied properties
    *
    * @param properties A plain object of property names and values to update the resource plan
    */
    public update = this._update<CommandResult<QueueJob>, TypedHash<any>, any>(
        "PS.ResourcePlan",
        data => ({ data: data, instance: this.getParent(DraftProject).queueJobs.getById(data.Id) }));

    /**
     * Deletes the ResourcePlan object
     */
    public async delete(): Promise<CommandResult<QueueJob>> {
        const data = await this.postCore({
            headers: {
                "X-HTTP-Method": "DELETE",
            },
        });
        return { data: data, instance: this.getParent(DraftProject).queueJobs.getById(data.Id) };
    }

    /**
     * Checks in the project resource plan in case it was left checked out in Project Server
     */
    public async forceCheckIn(): Promise<CommandResult<QueueJob>> {
        const data = await this.clone(ResourcePlan, "ForceCheckIn").postCore();
        return { data: data, instance: this.getParent(DraftProject).queueJobs.getById(data.Id) };
    }

    /**
     * Publishes the project resource plan and makes it visible to other users in Project Server
     */
    public async publish(): Promise<CommandResult<QueueJob>> {
        const data = await this.clone(ResourcePlan, "Publish").postCore();
        return { data: data, instance: this.getParent(DraftProject).queueJobs.getById(data.Id) };
    }
}
