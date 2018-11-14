import { ProjectHttpClient } from "./projecthttpclient";
import { combine, extend } from "@pnp/common";
import { ODataDefaultParser } from "@pnp/odata";
import { ProjectRuntimeConfig } from "../config/projectlibconfig";

export class CachedDigest {
    public expiration: Date;
    public value: string;
}

// allows for the caching of digests across all HttpClient's which each have their own DigestCache wrapper.
const digests = new Map<string, CachedDigest>();

export class DigestCache {

    constructor(private _httpClient: ProjectHttpClient, private _digests: Map<string, CachedDigest> = digests) { }

    public getDigest(webUrl: string): Promise<string> {

        const cachedDigest: CachedDigest = this._digests.get(webUrl);
        if (cachedDigest !== undefined) {
            const now = new Date();
            if (now < cachedDigest.expiration) {
                return Promise.resolve(cachedDigest.value);
            }
        }

        const url = combine(webUrl, "/_api/contextinfo");

        const headers = {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose;charset=utf-8",
        };

        return this._httpClient.fetchRaw(url, {
            cache: "no-cache",
            credentials: "same-origin",
            headers: extend(headers, ProjectRuntimeConfig.headers, true),
            method: "POST",
        }).then((response: any) => {
            const parser = new ODataDefaultParser();
            return parser.parse(response).then((d: any) => d.GetContextWebInformation);
        }).then((data: any) => {
            const newCachedDigest = new CachedDigest();
            newCachedDigest.value = data.FormDigestValue;
            const seconds = data.FormDigestTimeoutSeconds;
            const expiration = new Date();
            expiration.setTime(expiration.getTime() + 1000 * seconds);
            newCachedDigest.expiration = expiration;
            this._digests.set(webUrl, newCachedDigest);
            return newCachedDigest.value;
        });
    }

    public clear() {
        this._digests.clear();
    }
}
