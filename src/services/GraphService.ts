import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClientV3,  } from "@microsoft/sp-http";
import { ISite } from "../model/ISite";
import { IApp } from "../model/IApp";

export default class GraphService {
	private msGraphClientFactory: MSGraphClientFactory;
  private client: MSGraphClientV3;

  public static readonly serviceKey: ServiceKey<GraphService> =
    ServiceKey.create<GraphService>('react-application-nav-graph', GraphService);

  constructor(serviceScope: ServiceScope) {  
    serviceScope.whenFinished(async () => {
      this.msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);      
    });
  }

  public async searchSites(queryText: string, start: number): Promise<ISite[]> {
    const searchRawSites = await this.searchRawSites(queryText, start);
    const sites = await this.transformSearchSites(searchRawSites);
    return sites;
  }

  public async servicePrincipals(): Promise<IApp[]> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const response = await this.client
            .api(`servicePrincipals`)
            .version('v1.0')
            .filter(`startswith(DisplayName,'dlgScope')`)  // Assumption
            .get();
    
    const apps: IApp[] = [];
    response.value.forEach((app: any) => {
      apps.push({Id: app.appId, DisplayName: app.displayName });
    });
    return apps;
  }

  public async servicePrincipal(appId: string): Promise<string> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const response = await this.client
            .api(`servicePrincipals`)
            .version('v1.0')
            .filter(`appId eq '${appId}'`)
            .get();
    return response.value[0].displayName;
  }

  public async isSiteAdmin(userEMail: string, currentSiteId: string): Promise<boolean> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const response = await this.client
            .api(`sites/${currentSiteId}/lists/User Information List/items`)
            .version('v1.0')
            .header('Prefer','HonorNonIndexedQueriesWarningMayFailRandomly')
            .expand('fields($select=EMail,IsSiteAdmin)')
            .filter(`fields/EMail eq '${userEMail}'`)
            .get();
    return response.value[0].fields.IsSiteAdmin;
  }

  public async grantPermissions(role: string, appId: string, displayName: string, siteId: string): Promise<any[]> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const requestBody = {
      roles: [
        role
      ],
      grantedToIdentities: [
        {
          application: {
            id: appId,
            displayName: displayName
          }
        }
      ]
    };

    const response = await this.client
            .api(`/sites/${siteId}/permissions`)
            .version('v1.0')    
            .post(requestBody);
    return response;
  }

  private async searchRawSites(queryText: string, start: number): Promise<any[]> {
    this.client = await this.msGraphClientFactory.getClient('3');
    const requestBody = {
      requests: [
          {
              entityTypes: [
                  "site"
              ],
              query: {
                  "queryString": `${queryText}`
              }
          }
      ]
    };

    const response = await this.client
            .api(`search/query`)
            .version('v1.0')
            .skip(start)
            .top(20)   // Limit in batching!      
            .post(requestBody);
    if (response.value[0].hitsContainers[0].total > 0) {
      return response.value[0].hitsContainers[0].hits;
    }
    else return [];
  }

  private transformSearchSites(response: any[]): ISite[] {    
    const items: Array<ISite> = new Array<ISite>();
    if (response !== null && response.length > 0) {
      response.forEach((r: any) => {          
        items.push({ Title: r.resource.displayName, Url: r.resource.webUrl, Id: r.resource.id });
      });
      return items;
    }
    else {
      return [];
    }
  }
}