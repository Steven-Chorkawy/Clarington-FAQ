import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp, pnp logging system, and any other selective imports needed
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";

var _sp: SPFI = null;
var _siteSP: SPFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context != null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp;
};

export const getSiteSP = (context?: WebPartContext, siteUrl?: string): SPFI => {
  if (context != null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _siteSP = spfi(siteUrl).using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _siteSP;
}