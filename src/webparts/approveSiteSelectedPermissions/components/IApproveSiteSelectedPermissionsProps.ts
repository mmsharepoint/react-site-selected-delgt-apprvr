import { ServiceScope } from "@microsoft/sp-core-library";
// import { MSGraphClientFactory } from "@microsoft/sp-http";

export interface IApproveSiteSelectedPermissionsProps {
  description: string;
  isAdminMode: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteId: string;
  serviceScope: ServiceScope;
  selectedApp: string;
}
