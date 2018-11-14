import { Logger, LogLevel } from "@pnp/logging";
import { project } from "@pnp/project";
import { PnpNode } from "sp-pnp-node";

declare var process: { exit(code?: number): void };

export function Example(settings: any) {

    // configure your node options
    project.setup({
        project: {
            baseUrl: settings.testing.project.url,
            fetchClientFactory: () => new PnpNode({
                authOptions: settings.testing.project,
                siteUrl: settings.testing.project.url,
            }),
        },
    });

    // run some debugging
    project.projects.add({ name: "TestProject" }).then(data => {

        // logging results to the Logger
        Logger.log({
            data: JSON.stringify(data),
            level: LogLevel.Info,
            message: "Result",
        });

        process.exit(0);
    }).catch(e => {

        Logger.error(e);
        process.exit(1);
    });
}
