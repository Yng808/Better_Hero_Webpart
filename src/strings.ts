import { ContextInfo } from "gd-sprest-bs";

// Sets the context information
// This is for SPFx or Teams solutions
export const setContext = (context, sourceUrl?: string) => {
  // Set the context
  ContextInfo.setPageContext(context.pageContext);

  // Update the source url
  Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
};

/**
 * Global Constants
 */
const Strings = {
  AppElementId: "BetterHero",
  GlobalVariable: "BetterHero",
  /*
    Lists: {
        Main: "Dashboard",
        Devices: "Devices"
    },
    */
  Libraries: {
    Main: "BetterHeroLibrary",
  },
  ProjectName: "Better Hero Webpart",
  ProjectDescription: "Configurable hero web part using bootstrap cards",
  SourceUrl: ContextInfo.webServerRelativeUrl,
  Version: "0.1",
};
export default Strings;
