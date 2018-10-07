import { Logger, LogLevel } from "@pnp/logging";
import { sp } from "@pnp/sp";
import { project } from "@pnp/project";
import { PnpNode } from "sp-pnp-node";

declare var process: { exit(code?: number): void };

export function Example(settings: any) {

    // configure your node options
    sp.setup({
        sp: {
            baseUrl: settings.testing.project.url,
            fetchClientFactory: () => new PnpNode({
                authOptions: settings.testing.project,
                siteUrl: settings.testing.project.url,
            }),
        },
    });

    // run some debugging
    project.projects.select("Name").get().then(projects => {

        // logging results to the Logger
        projects.forEach(p =>
            Logger.log({
                data: p.Name,
                level: LogLevel.Info,
                message: "Project Name",
            }));

        process.exit(0);
    }).catch(e => {

        Logger.error(e);
        process.exit(1);
    });
}
