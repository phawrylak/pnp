import { jsS, TypedHash } from "@pnp/common";
import { wrap } from "./utils/wrap";
import {
    defaultPath,
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";
import {
    ProjectDetailPageCollection,
    ProjectDetailPageCreationInformation,
} from "./projectdetailpages";

/**
 * Represents a collection of EnterpriseProjectType (EPT) objects
 */
@defaultPath("_api/ProjectServer/EnterpriseProjectTypes")
export class EnterpriseProjectTypeCollection extends ProjectQueryableCollection {

    /**
    * Gets an enterprise project type (EPT) from the collection with the specified GUID
    *
    * @param id The string representation of the EPT GUID
    */
    public getById(id: string): EnterpriseProjectType {
        const ept = new EnterpriseProjectType(this);
        ept.concat(`('${id}')`);
        return ept;
    }

    /**
     * Adds the enterprise project type (EPT) that is specified by the EnterpriseProjectTypeCreationInformation object to the collection
     *
     * @param parameters The properties of the EPT to create
     */
    public async add(parameters: EnterpriseProjectTypeCreationInformation): Promise<CommandResult<EnterpriseProjectType>> {
        const data = await this.clone(EnterpriseProjectTypeCollection, "Add").postCore({ body: jsS(wrap({parameters: parameters})) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Creates an object that represents an enterprise project type
 */
export class EnterpriseProjectType extends ProjectQueryableInstance {

    /**
     * Gets the collection of enterprise project detail pages
     */
    public get projectDetailPages(): ProjectDetailPageCollection {
        return new ProjectDetailPageCollection(this, "ProjectDetailPages");
    }

    /**
     * Deletes the current EnterpriseProjectType object
     */
    public delete = this._delete;

    /**
    * Updates this enterprise project type with the supplied properties
    *
    * @param properties A plain object of property names and values to update the enterprise project type
    */
    public update = this._update<void, TypedHash<any>, any>(
        "PS.EnterpriseProjectType",
        _ => null);
}

/**
 * Provides information for the creation of an enterprise project type (EPT)
 */
export interface EnterpriseProjectTypeCreationInformation {

    /**
     * Gets or sets the GUIDs of departments that are associated with an EPT
     */
    DepartmentIds?: string[];

    /**
     * Gets or sets the description for an EPT
     */
    Description?: string;

    /**
     * Gets or sets the GUID for an EPT
     */
    Id?: string;

    /**
     * Gets or sets the URL of an image that is associated with an EPT
     */
    ImageUrl?: string;

    /**
     * Gets or sets whether an EPT is the type that all new projects should use by default
     */
    IsDefault?: boolean;

    /**
     * Gets or sets whether a project that an EPT creates is fully managed by Project Server, or is a SharePoint tasks list
     */
    IsManaged?: boolean;

    /**
     * Gets or sets the name of an EPT
     */
    Name: string;

    /**
     * Gets or sets a value that indicates whether the user wants the enterprise project type to appear at the end of the list of EPTs,
     * or whether the user wants to control where it is placed in the list
     */
    Order?: number;

    /**
     * Gets or sets whether a project that an EPT creates has permissions synchronization enabled
     */
    PermissionSyncEnable?: boolean;

    /**
     * Gets or sets the project detail page that is used as the first page in the workflow for an enterprise project type
     */
    ProjectDetailPages: ProjectDetailPageCreationInformation[];

    /**
     * Gets or sets the GUID of the project plan template that was created with an EPT
     */
    ProjectPlanTemplateId?: string;

    /**
     * Gets or sets the site creation option for an EPT
     */
    SiteCreationOption?: EnterpriseProjectTypeSiteCreationOptions;

    /**
     * Gets or sets the site creation URL for an EPT
     */
    SiteCreationURL?: string;

    /**
     * Gets or sets whether a project that an EPT creates has task list synchronization enabled
     */
    TaskListSyncEnable?: boolean;

    /**
     * Gets or sets the GUID of the workflow that is associated with an EPT
     */
    WorkflowAssociationId?: string;

    /**
     * Gets or sets the name of the workflow that is associated with an EPT
     */
    WorkflowAssociationName?: string;

    /**
     * Gets or sets the LCID of the project site template that is associated with an EPT
     */
    WorkspaceTemplateLCID?: string;

    /**
     * Gets or sets the name of the project site template that is associated with an EPT
     */
    WorkspaceTemplateName: string;
}

/**
 * Specifies how project sites should be created for an enterprise project type (EPT)
 */
export enum EnterpriseProjectTypeSiteCreationOptions {

    /**
     * The project site creation is not specified for the EPT
     */
    NotSpecified,

    /**
     * Allow users to choose on publish whether project site should be created
     */
    AskOnPublish,

    /**
     * Create project site on first publish
     */
    CreateOnFirstPublish,

    /**
     * Do not create a project site
     */
    None,
}
