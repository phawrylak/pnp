import { ConfigOptions } from "@pnp/common";
import {
    ProjectQueryable,
    ProjectQueryableConstructor,
} from "./projectqueryable";
import { CalendarCollection } from "./calendars";
import { CustomFieldCollection } from "./customfields";
import { EnterpriseProjectTypeCollection } from "./enterpriseprojecttypes";
import { EnterpriseResourceCollection } from "./enterpriseresources";
import { EntityTypes } from "./entitytypes";
import { LookupTableCollection } from "./lookuptables";
import { PhaseCollection } from "./phases";
import { ProjectDetailPageCollection } from "./projectdetailpages";
import { ProjectCollection } from "./projects";
import { JobState, QueueJob } from "./queuejobs";
import { StageCollection } from "./stages";
import { TimeSheetPeriodCollection } from "./timesheetperiods";

/**
 * Root of the Project REST module
 */
export class ProjectRest {

    /**
     * Creates a new instance of the ProjectRest class
     *
     * @param options Additional options
     * @param baseUrl A string that should form the base part of the url
     */
    constructor(protected _options: ConfigOptions = {}, protected _baseUrl = "") { }

    /**
     * TODO
     */
    public get calendars(): CalendarCollection {
        return this.create(CalendarCollection);
    }

    /**
     * TODO
     */
    public get customFields(): CustomFieldCollection {
        return this.create(CustomFieldCollection);
    }

    /**
     * TODO
     */
    public get enterpriseProjectTypes(): EnterpriseProjectTypeCollection {
        return this.create(EnterpriseProjectTypeCollection);
    }

    /**
     * TODO
     */
    public get enterpriseResources(): EnterpriseResourceCollection {
        return this.create(EnterpriseResourceCollection);
    }

    /**
     * TODO
     */
    public get entityTypes(): EntityTypes {
        return this.create(EntityTypes);
    }

    /**
     * TODO
     */
    public get lookupTables(): LookupTableCollection {
        return this.create(LookupTableCollection);
    }

    /**
     * TODO
     */
    public get phases(): PhaseCollection {
        return this.create(PhaseCollection);
    }

    /**
     * TODO
     */
    public get projectDetailPages(): ProjectDetailPageCollection {
        return this.create(ProjectDetailPageCollection);
    }

    /**
     * Gets projects
     */
    public get projects(): ProjectCollection {
        return this.create(ProjectCollection);
    }

    /**
     * TODO
     */
    public get stages(): StageCollection {
        return this.create(StageCollection);
    }

    /**
     * TODO
     */
    public get timeSheetPeriods(): TimeSheetPeriodCollection {
        return this.create(TimeSheetPeriodCollection);
    }

    /**
     * Configures instance with additional options and baseUrl.
     * Provided configuration used by other objects in a chain
     *
     * @param options Additional options
     * @param baseUrl A string that should form the base part of the url
     */
    public configure(options: ConfigOptions, baseUrl = ""): ProjectRest {
        return new ProjectRest(options, baseUrl);
    }

    /**
     * Waits for the specified queue job to complete, or for a maximum number of seconds
     *
     * @param job An object that represents the queued job
     * @param timeoutSeconds The maximum number of seconds to wait for the queue job to complete
     */
    public async waitForQueue(job: QueueJob, timeoutSeconds: number): Promise<JobState> {
        const waitingStates = [
            JobState.Unknown,
            JobState.ReadyForProcessing,
            JobState.SendIncomplete,
            JobState.Processing,
            JobState.ProcessingDeferred,
            JobState.OnHold,
            JobState.Sleeping,
            JobState.ReadyForLaunch,
        ];

        do {
            const data = await job.get();
            if (!data) {
                return JobState.Success;
            } else if (waitingStates.indexOf(data.JobState) === -1) {
                return data.JobState;
            } else {
                await new Promise(resolve => setTimeout(resolve, 2000));
                timeoutSeconds -= 2;
            }
        } while (timeoutSeconds > 0);
    }

    /**
     * Handles creating and configuring the objects returned from this class
     *
     * @param fm The factory method used to create the instance
     * @param path Optional additional path information to pass to the factory method
     */
    private create<T extends ProjectQueryable>(fm: ProjectQueryableConstructor<T>, path?: string): T {
        return new fm(this._baseUrl, path).configure(this._options);
    }
}

export const project = new ProjectRest();
