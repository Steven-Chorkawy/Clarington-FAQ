export interface IFaqAccordionWebPartProps {
  description: string;
  siteUrl: string;        // URL of the SharePoint site.
  listName: string;       // Name of the SharePoint list.
  questionFieldName: string;  // Name of the Question field.
  answerFieldName: string;    // Name of the Answer field. 
}

export interface IFaqAccordionProps extends IFaqAccordionWebPartProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
