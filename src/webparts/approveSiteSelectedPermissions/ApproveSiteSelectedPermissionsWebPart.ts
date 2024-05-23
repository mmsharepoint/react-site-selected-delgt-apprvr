import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  // PropertyPaneTextField,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ApproveSiteSelectedPermissionsWebPartStrings';
import { ApproveSiteSelectedPermissions } from './components/ApproveSiteSelectedPermissions';
import { IApproveSiteSelectedPermissionsProps } from './components/IApproveSiteSelectedPermissionsProps';
import GraphService from '../../services/GraphService';
import { IApp } from '../../model/IApp';

export interface IApproveSiteSelectedPermissionsWebPartProps {
  isAdminMode: boolean;
  selectedApp: string;
}

export default class ApproveSiteSelectedPermissionsWebPart extends BaseClientSideWebPart<IApproveSiteSelectedPermissionsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _servicePrincipals: IApp[];

  public render(): void {
    const element: React.ReactElement<IApproveSiteSelectedPermissionsProps> = React.createElement(
      ApproveSiteSelectedPermissions,
      {
        isAdminMode: this.properties.isAdminMode,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        serviceScope: this.context.serviceScope,
        selectedApp: this.properties.selectedApp,
        siteId: this.context.pageContext.site.id + ',' + this.context.pageContext.web.id
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // https://www.voitanos.io/blog/sharepoint-framework-dynamic-property-pane-dropdown/
    const graphService = new GraphService(this.context.serviceScope);

    this._servicePrincipals = await graphService.servicePrincipals();
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                /* PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }), */
                PropertyPaneToggle('isAdminMode', {
                  label: strings.PropertyPaneIsAdminMode
                }),
                PropertyPaneDropdown('selectedApp', {
                  label: 'Service Principal',
                  options: this._servicePrincipals.map((app: IApp) => {
                    return <IPropertyPaneDropdownOption>{
                      key: app.Id, text: app.DisplayName
                    }
                  })
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
