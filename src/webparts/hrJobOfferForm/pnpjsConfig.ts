import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";

let _sp: SPFI = null;

export const getSP = (context?: WebPartContext) : SPFI => {
    if(_sp === null && context !== null) {
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};