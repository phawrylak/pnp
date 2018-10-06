import { ConfigOptions } from "@pnp/common";

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
     * Configures instance with additional options and baseUrl.
     * Provided configuration used by other objects in a chain
     * 
     * @param options Additional options
     * @param baseUrl A string that should form the base part of the url
     */
    public configure(options: ConfigOptions, baseUrl = ""): ProjectRest {
        return new ProjectRest(options, baseUrl);
    }
}

export const project = new ProjectRest();
