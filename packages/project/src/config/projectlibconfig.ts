import {
    LibraryConfiguration,
    TypedHash,
    RuntimeConfig,
    HttpClientImpl,
    FetchClient,
} from "@pnp/common";

export interface ProjectConfigurationPart {
    project?: {
        /**
         * Any headers to apply to all requests
         */
        headers?: TypedHash<string>;

        /**
         * The base url used for all requests
         */
        baseUrl?: string;

        /**
         * Defines a factory method used to create fetch clients
         */
        fetchClientFactory?: () => HttpClientImpl;
    };
}

export interface ProjectConfiguration extends LibraryConfiguration, ProjectConfigurationPart { }

export function setup(config: ProjectConfiguration): void {
    RuntimeConfig.extend(config);
}

export class ProjectRuntimeConfigImpl {

    public get headers(): TypedHash<string> {

        const projectPart = RuntimeConfig.get("project");
        if (projectPart !== undefined && projectPart.headers !== undefined) {
            return projectPart.headers;
        }

        return {};
    }

    public get baseUrl(): string | null {

        const projectPart = RuntimeConfig.get("project");
        if (projectPart !== undefined && projectPart.baseUrl !== undefined) {
            return projectPart.baseUrl;
        }

        if (RuntimeConfig.spfxContext !== undefined && RuntimeConfig.spfxContext !== null) {
            return RuntimeConfig.spfxContext.pageContext.web.absoluteUrl;
        }

        return null;
    }

    public get fetchClientFactory(): () => HttpClientImpl {

        const projectPart = RuntimeConfig.get("project");
        if (projectPart !== undefined && projectPart.fetchClientFactory !== undefined) {
            return projectPart.fetchClientFactory;
        } else {
            return () => new FetchClient();
        }
    }
}

export let ProjectRuntimeConfig = new ProjectRuntimeConfigImpl();
