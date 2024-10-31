import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FaqAccordionWebPartStrings';
import FaqAccordion from './components/FaqAccordion';
import { IFaqAccordionProps, IFaqAccordionWebPartProps } from './components/IFaqAccordionProps';
import { getSP, getSiteSP } from '../../pnpjs-config';
import "@pnp/sp/sites";
import Loading from './components/Loading';
import '@progress/kendo-theme-fluent/dist/all.css'

export default class FaqAccordionWebPart extends BaseClientSideWebPart<IFaqAccordionWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IFaqAccordionProps> = React.createElement(
      FaqAccordion,
      {
        description: this.properties.description,
        siteUrl: this.properties.siteUrl,
        listName: this.properties.listName,
        questionFieldName: this.properties.questionFieldName,
        answerFieldName: this.properties.answerFieldName,
        subtitleFieldName: this.properties.subtitleFieldName,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    const loadingElement: React.ReactElement<any> = React.createElement(Loading);

    if (this.properties.siteUrl && this.properties.listName)
      ReactDom.render(element, this.domElement);
    else
      ReactDom.render(loadingElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = await this._getEnvironmentMessage();

    super.onInit();

    //Initialize our _sp object that we can then use in other packages without having to pass around the context.
    //  Check out pnpjsConfig.ts for an example of a project setup file.
    getSP(this.context);

    if (this.properties.siteUrl) {
      getSiteSP(this.context, this.properties.siteUrl);
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _validatePropertyPaneList(listName: string): Promise<string> {
    const listURL = `${this.properties.siteUrl}/Lists/${listName}`;
    const errorMessage = `Cannot locate the list '${listURL}'...`;
    try {
      // This will throw an error if the list does not exist.       
      await getSiteSP().web.lists.getByTitle(listName)();

      return "";
    } catch (error) {
      console.error(error);
      return errorMessage;
    }
  }

  private async _validatePropertyPaneSite(siteUrl: string): Promise<string> {
    const errorMessage = `Cannot locate Site '${siteUrl}'...`
    try {
      const site = await getSP().site.exists(siteUrl);
      debugger;
      if (site) {
        getSiteSP(this.context, siteUrl);
        return '';
      }
      else {
        return errorMessage;
      }
    } catch (error) {
      console.error(error);
      return errorMessage;
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description"
                }),
                PropertyPaneTextField('siteUrl', {
                  label: "Site URL",
                  onGetErrorMessage: this._validatePropertyPaneSite.bind(this),
                }),
                PropertyPaneTextField('listName', {
                  label: "List Name",
                  onGetErrorMessage: this._validatePropertyPaneList.bind(this),
                }),
                PropertyPaneTextField('questionFieldName', {
                  label: "Question Field Name"
                }),
                PropertyPaneTextField('answerFieldName', {
                  label: "Answer Field Name"
                }),
                PropertyPaneTextField('subtitleFieldName', {
                  label: "Subtitle Field Name"
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
