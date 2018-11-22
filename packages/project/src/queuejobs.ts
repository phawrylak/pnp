import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { DraftProject, Project } from "./projects";
import { User } from "./users";

/**
 * Represents a collection of QueueJob objects
 */
export class QueueJobCollection extends ProjectQueryableCollection {

    /**
    * Gets a queue job from the collection with the specified GUID
    *
    * @param id The string representation of the queue job GUID
    */
    public getById(id: string): QueueJob {
        const job = new QueueJob(this);
        job.concat(`('${id}')`);
        return job;
    }
}

/**
 * Queues a project for publishing
 */
export class QueueJob extends ProjectQueryableInstance {

    /**
     * Gets the project that is queued
     */
    public get project(): Project {
        return new DraftProject(this, "Project");
    }

    /**
     * Gets the resource that submitted the queue job
     */
    public get submitter(): User {
        return new User(this, "Submitter");
    }

    /**
     * Cancels the queue job
     */
    public cancel(): Promise<void> {
        return this.clone(QueueJob, "Cancel").postCore();
    }
}

/**
 * The Project Server queue job state specifies the status of a queue job
 */
export enum JobState {

    /**
     * The queue job is unknown
     */
    Unknown,

    /**
     * The queue job is ready for processing
     */
    ReadyForProcessing,

    /**
     * The message sent to the queue is incomplete
     */
    SendIncomplete,

    /**
     * The queue job is processing
     */
    Processing,

    /**
     * The queue job completed successfully
     */
    Success,

    /**
     * The queue job failed. Unfinished jobs with the same correlation are blocked
     */
    Failed,

    /**
     * The queue job failed. Unfinished jobs with the same correlation are not blocked
     */
    FailedNotBlocking,

    /**
     * The queue job processing is deferred
     */
    ProcessingDeferred,

    /**
     * The queue job correlation is blocked; the job is not processed
     */
    CorrelationBlocked,

    /**
     * The queue job is canceled
     */
    Canceled,

    /**
     * The queue job is on hold
     */
    OnHold,

    /**
     * The queue job is sleeping
     */
    Sleeping,

    /**
     * The queue job is ready for launch
     */
    ReadyForLaunch,

    /**
     * This is the last state of the queue job (internal use only)
     */
    LastState,
}
