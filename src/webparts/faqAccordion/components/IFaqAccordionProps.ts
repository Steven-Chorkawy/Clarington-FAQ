export interface IFaqAccordionProps {
  //#region Props from property pane
  description: string;
  siteUrl: string;      // URL of the SharePoint site.
  listName: string;     // Name of the SharePoint list.
  //#endregion
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
