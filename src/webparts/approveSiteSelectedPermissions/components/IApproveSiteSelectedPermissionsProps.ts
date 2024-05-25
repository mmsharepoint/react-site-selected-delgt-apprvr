import { ServiceScope } from "@microsoft/sp-core-library";

export interface IApproveSiteSelectedPermissionsProps {
  hasTeamsContext: boolean;
  isAdminMode: boolean;
  userEMail: string;
  userDisplayName: string;
  siteId: string;
  serviceScope: ServiceScope;
  selectedApp: string;
}
