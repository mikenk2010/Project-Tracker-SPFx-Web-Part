import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const getSP = (context: WebPartContext): SPFI => {
  if (!context) {
    throw new Error("SPFx context is required to initialize PnPjs");
  }
  
  try {
    // Create a new instance each time to ensure context is fresh
    // SPFx(context) sets up the authentication and request context
    const sp = spfi().using(SPFx(context));
    
    if (!sp) {
      throw new Error("Failed to initialize PnPjs - spfi returned undefined");
    }
    
    // Verify web property exists
    if (!(sp as any).web) {
      console.error("PnPjs initialized but web property is missing");
      console.error("SP object keys:", Object.keys(sp));
      throw new Error("PnPjs web property is missing after initialization");
    }
    
    return sp;
  } catch (error: any) {
    console.error("Error initializing PnPjs:", error);
    throw new Error(`Failed to initialize PnPjs: ${error.message || error}`);
  }
};

