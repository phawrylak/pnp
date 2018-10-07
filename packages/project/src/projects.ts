import { defaultPath, SharePointQueryableCollection } from "@pnp/sp";

import { Project } from "./project";

/**
 * Describes a collection of projects
 *
 */
@defaultPath("_api/ProjectServer/Projects")
export class Projects extends SharePointQueryableCollection {

    /**
    * Gets a Project by id
    *
    * @param id The GUID of the project to retrieve
    */
    public getById(id: string): Project {
        const p = new Project(this);
        p.concat(`('${id}')`);
        return p;
    }
}
