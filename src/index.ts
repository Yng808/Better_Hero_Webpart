import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { InstallationRequired, LoadingDialog } from "dattatable";
import { Configuration } from "./cfg";
import Strings, { setContext } from "./strings";

// create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el, context?, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Update the configuration
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);
        }
    }
}