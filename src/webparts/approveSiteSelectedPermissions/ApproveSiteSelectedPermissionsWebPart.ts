import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ApproveSiteSelectedPermissionsWebPartStrings';
import { ApproveSiteSelectedPermissions } from './components/ApproveSiteSelectedPermissions';
import { IApproveSiteSelectedPermissionsProps } from './components/IApproveSiteSelectedPermissionsProps';
import { IApp } from '../../model/IApp';
import FunctionService from '../../services/FunctionService';

export interface IApproveSiteSelectedPermissionsWebPartProps {
  isAdminMode: boolean;
  selectedApp: string;
}

export default class ApproveSiteSelectedPermissionsWebPart extends BaseClientSideWebPart<IApproveSiteSelectedPermissionsWebPartProps> {
  private _servicePrincipals: IApp[];

  public render(): void {
    const absoluteUrl: URL = new URL(this.context.pageContext.site.absoluteUrl);
    const host = absoluteUrl.hostname;
    const element: React.ReactElement<IApproveSiteSelectedPermissionsProps> = React.createElement(
      ApproveSiteSelectedPermissions,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        isAdminMode: this.properties.isAdminMode,
        userEMail: this.context.pageContext.user.email,
        userDisplayName: this.context.pageContext.user.displayName,
        serviceScope: this.context.serviceScope,
        selectedApp: this.properties.selectedApp,
        site: {
          Id: host + ',' + this.context.pageContext.site.id + ',' + this.context.pageContext.web.id,
          Url: this.context.pageContext.site.absoluteUrl,
          Title: ''
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    // https://www.voitanos.io/blog/sharepoint-framework-dynamic-property-pane-dropdown/
    const functionService = new FunctionService(this.context.serviceScope);
    this._servicePrincipals = await functionService.servicePrincipals('dlg');
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

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
