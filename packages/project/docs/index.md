# @pnp/project

[![npm version](https://badge.fury.io/js/%40pnp%2Fproject.svg)](https://badge.fury.io/js/%40pnp%2Fproject)

This package contains the fluent API used to call the Project REST services.

## Getting Started

Install the library and required dependencies

`npm install @pnp/logging @pnp/common @pnp/odata @pnp/project --save`

Import the library into your application and access the root project object

```TypeScript
import { project } from "@pnp/project";

(function main() {

    // here we will load list of projects
    project.projects.select("Name").get().then(projects => {
        projects.forEach(p => console.log(`Project Name: ${p.Name}`));
    });
})()
```

## Getting Started: SharePoint Framework

Install the library and required dependencies

`npm install @pnp/logging @pnp/common @pnp/odata @pnp/project --save`

Import the library into your application, update OnInit, and access the root project object in render

```TypeScript
import { project } from "@pnp/project";

// ...

public onInit(): Promise<void> {

  return super.onInit().then(_ => {

    // other init code may be present

    project.setup({
      spfxContext: this.context
    });
  });
}

// ...

public render(): void {

    // A simple loading message
    this.domElement.innerHTML = `Loading...`;

    project.projects.select("Name").get().then(projects => {
        this.domElement.innerHTML = projects.map(p => p.Name).join(', ');
    });
}
```

## Getting Started: Node.js

Install the library and required dependencies

`npm install @pnp/logging @pnp/common @pnp/odata @pnp/project @pnp/nodejs --save`

Import the library into your application, setup the node client, make a request

```TypeScript
import { project } from "@pnp/project";
import { SPFetchClient } from "@pnp/nodejs";

// do this once
project.setup({
    project: {
        fetchClientFactory: () => {
            return new SPFetchClient("{your site url}", "{your client id}", "{your client secret}");
        },
    },
});

// now make any calls you need using the configured client
project.projects.select("Name").get().then(projects => {
    projects.forEach(p => console.log(`Project Name: ${p.Name}`));
});
```

## Library Topics

* [Projects](projects.md)

## UML
TODO