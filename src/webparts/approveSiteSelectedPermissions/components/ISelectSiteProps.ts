import { ServiceScope } from "@microsoft/sp-core-library";

export interface ISelectSiteProps {
  serviceScope: ServiceScope;
  siteSelectedCallback: (s: string) => void;
}