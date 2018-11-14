import { ConfigOptions } from "@pnp/common";
import {
    ProjectQueryable,
    ProjectQueryableConstructor,
} from "./projectqueryable";
import { Projects } from "./projects";

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

    /**
     * Gets projects
     */
    public get projects(): Projects {
        return this.create(Projects);
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
