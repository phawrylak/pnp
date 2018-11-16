import { jsS } from "@pnp/common";
import {
    ProjectQueryableCollection,
    ProjectQueryableInstance,
} from "./projectqueryable";
import { CommandResult } from "./types";

/**
 * Represents a collection of EnterpriseProjectType (EPT) objects
 */
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
        const data = await this.postCore({ body: jsS(parameters) });
        return { data: data, instance: this.getById(data.Id) };
    }
}

/**
 * Creates an object that represents an enterprise project type
 */
export class EnterpriseProjectType extends ProjectQueryableInstance {

    /**
     * Deletes the current EnterpriseProjectType object
     */
    public delete = this._delete;
}

/**
 * Provides information for the creation of an enterprise project type (EPT)
 */
export interface EnterpriseProjectTypeCreationInformation {

    /**
     * Gets or sets the GUIDs of departments that are associated with an EPT
     */
    departmentIds?: string[];

    /**
     * Gets or sets the description for an EPT
     */
    description?: string;

    /**
     * Gets or sets the GUID for an EPT
     */
    id?: string;

    /**
     * Gets or sets the URL of an image that is associated with an EPT
     */
    imageUrl?: string;

    /**
     * Gets or sets whether an EPT is the type that all new projects should use by default
     */
    isDefault?: boolean;

    /**
     * Gets or sets whether a project that an EPT creates is fully managed by Project Server, or is a SharePoint tasks list
     */
    isManaged?: boolean;

    /**
     * Gets or sets the name of an EPT
     */
    name: string;

    /**
     * Gets or sets a value that indicates whether the user wants the enterprise project type to appear at the end of the list of EPTs,
     * or whether the user wants to control where it is placed in the list
     */
    order?: number;

    /**
     * Gets or sets whether a project that an EPT creates has permissions synchronization enabled
     */
    permissionSyncEnable?: boolean;

    /**
     * Gets or sets the project detail page that is used as the first page in the workflow for an enterprise project type
     */
    projectDetailPages?: any[]; // TODO: ProjectDetailPageCreationInformation[]

    /**
     * Gets or sets the GUID of the project plan template that was created with an EPT
     */
    projectPlanTemplateId?: string;

    /**
     * Gets or sets the site creation option for an EPT
     */
    siteCreationOption?: number; // TODO: enum?

    /**
     * Gets or sets the site creation URL for an EPT
     */
    siteCreationURL?: string;

    /**
     * Gets or sets whether a project that an EPT creates has task list synchronization enabled
     */
    taskListSyncEnable?: boolean;

    /**
     * Gets or sets the GUID of the workflow that is associated with an EPT
     */
    workflowAssociationId?: string;

    /**
     * Gets or sets the name of the workflow that is associated with an EPT
     */
    workflowAssociationName?: string;

    /**
     * Gets or sets the LCID of the project site template that is associated with an EPT
     */
    workspaceTemplateLCID?: number;

    /**
     * Gets or sets the name of the project site template that is associated with an EPT
     */
    workspaceTemplateName?: string;
}
