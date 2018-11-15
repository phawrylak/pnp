import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";

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
     * Cancels the queue job
     */
    public cancel(): Promise<void> {
        return this.clone(QueueJob, "Cancel").postCore();
    }

    /**
     * Waits until the job is finished (uses polling)
     *
     * @param timeout The maximum time to wait (ms)
     * @param interval The polling interval (ms)
     */
    public async waitForJob(timeout = 0, interval = 250): Promise<QueueJobResult> {
        const finishedStates = [
            JobState.SendIncomplete,
            JobState.Success,
            JobState.Failed,
            JobState.FailedNotBlocking,
            JobState.CorrelationBlocked,
            JobState.Canceled,
            JobState.LastState,
        ];

        let isTimeout = false;
        if (timeout) {
            setTimeout(() => isTimeout = true, timeout);
        }

        while (true) {
            const data = await this.get();
            if (finishedStates.indexOf(data.JobState) > -1 || isTimeout) {
                return { data: data, queueJob: this };
            } else {
                await new Promise(resolve => setTimeout(resolve, interval));
            }
        }
    }
}

export interface QueueJobResult {
    data: any;
    queueJob: QueueJob;
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
