import { ServiceScope } from "@microsoft/sp-core-library";
import { ISite } from "../../../model/ISite";

export interface IApproveSiteSelectedPermissionsProps {
  hasTeamsContext: boolean;
  isAdminMode: boolean;
  userEMail: string;
  userDisplayName: string;
  site: ISite;
  serviceScope: ServiceScope;
  selectedApp: string;
}
