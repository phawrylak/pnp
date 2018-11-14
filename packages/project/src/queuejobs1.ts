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
}
