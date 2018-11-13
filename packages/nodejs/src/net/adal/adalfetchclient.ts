declare var require: (path: string) => any;
import { AuthenticationContext } from "adal-node";
const nodeFetch = require("node-fetch").default;

import {
    combine,
    objectDefinedNotNull,
    HttpClientImpl,
    isUrlAbsolute,
    extend,
} from "@pnp/common";

export interface AADToken {
    accessToken: string;
    expiresIn: number;
    expiresOn: string | Date;
    isMRRT?: boolean;
    resource: string;
    tokenType: string;
}

export class AdalFetchClient implements HttpClientImpl {

    private authContext: any;

    constructor(private _tenant: string,
        private _clientId: string,
        private _secret: string,
        private _resource = "https://graph.microsoft.com",
        private _authority = "https://login.windows.net") {

        this.authContext = new AuthenticationContext(combine(this._authority, this._tenant));
    }

    public fetch(url: string, options: any): Promise<Response> {

        if (!objectDefinedNotNull(options)) {
            options = {
                headers: new Headers(),
            };
        } else if (!objectDefinedNotNull(options.headers)) {
            options = extend(options, {
                headers: new Headers(),
            });
        }

        if (!isUrlAbsolute(url)) {
            url = combine(this._resource, url);
        }

        return this.acquireToken().then(token => {

            options.headers.set("Authorization", `${token.tokenType} ${token.accessToken}`);

            return nodeFetch(url, options);
        });
    }

    public acquireToken(): Promise<AADToken> {
        return new Promise((resolve, reject) => {

            this.authContext.acquireTokenWithClientCredentials(this._resource, this._clientId, this._secret, (err: any, token: AADToken) => {

                if (err) {
                    reject(err);
                } else {
                    resolve(token);
                }
            });
        });
    }
}
