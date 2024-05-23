import { ServiceScope } from "@microsoft/sp-core-library";

export interface IApproveSiteSelectedPermissionsProps {
  isAdminMode: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteId: string;
  serviceScope: ServiceScope;
  selectedApp: string;
}
