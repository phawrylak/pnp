import { hOP, isArray, jsS } from "@pnp/common";

/**
 * Wraps arrays to object with results property
 *
 * @param obj Object to process
 */
export function wrap(obj: any): any {
    const copy = JSON.parse(jsS(obj));
    processObject(copy);
    return copy;
}

function processObject(obj: any): void {
    for (const property in obj) {
        if (hOP(obj, property)) {
            if (isArray(obj[property])) {
                obj[property] = { results: obj[property] };
            } else if (obj[property] === Object(obj[property])) {
                processObject(obj[property]);
            }
        }
    }
}
