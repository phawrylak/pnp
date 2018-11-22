import { defaultPath, ProjectQueryableInstance } from "./projectqueryable";

/**
 * Represents the types of Project Server entities
 */
@defaultPath("_api/ProjectServer/EntityTypes")
export class EntityTypes extends ProjectQueryableInstance {

    /**
     * Gets an assignment entity type
     */
    public get assignmentEntity(): EntityType {
        return new EntityType(this, "AssignmentEntity");
    }

    /**
     * Gets a project entity type
     */
    public get projectEntity(): EntityType {
        return new EntityType(this, "ProjectEntity");
    }

    /**
     * Gets a resource entity type
     */
    public get resourceEntity(): EntityType {
        return new EntityType(this, "ResourceEntity");
    }

    /**
     * Gets a task entity type
     */
    public get taskEntity(): EntityType {
        return new EntityType(this, "TaskEntity");
    }
}

/**
 * Represents a type of Project Server entity
 */
export class EntityType extends ProjectQueryableInstance {
}
