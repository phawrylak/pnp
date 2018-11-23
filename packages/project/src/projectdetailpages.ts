import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";

/**
 * Represents a collection of project detail pages (PDPs)
 */
@defaultPath("_api/ProjectServer/ProjectDetailPages")
export class ProjectDetailPageCollection extends ProjectQueryableCollection {

    /**
    * Gets a project detail page from the list of available project detail pages by using a GUID identifier
    *
    * @param id The string representation of a project detail page GUID
    */
    public getById(id: string): ProjectDetailPage {
        const page = new ProjectDetailPage(this);
        page.concat(`('${id}')`);
        return page;
    }
}

/**
 * Represents a project detail page (PDP), which is a Web Part page for creating, viewing, or managing the properties of projects in Project Web App
 */
export class ProjectDetailPage extends ProjectQueryableInstance {
}

/**
 * Provides information that is used to create a project detail page (PDP) for an enterprise project type
 */
export interface ProjectDetailPageCreationInformation {

    /**
     * Gets or sets the GUID of a project detail page
     */
    Id: string;

    /**
     * Gets or sets a value that indicates whether a project detail page is created
     */
    IsCreate?: boolean;

    /**
     * Gets or sets a value that indicates the order of placement of a project detail page in a list of project detail pages
     */
    Position?: number;
}
