import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClientFactory, AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { ISite } from "../model/ISite";
import { IApp } from "../model/IApp";
const config: any = require('./azFunct.json');


export default class FunctionService {
	private aadHttpClientFactory: AadHttpClientFactory;
  private client: AadHttpClient;

  public static readonly serviceKey: ServiceKey<FunctionService> =
    ServiceKey.create<FunctionService>('react-site-selected-delgt-apprvr', FunctionService);

  constructor(serviceScope: ServiceScope) {  
    serviceScope.whenFinished(async () => {
      this.aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);      
    });
  }

  public async searchSites(queryText: string, start: number): Promise<ISite[]> {
    this.client = await this.aadHttpClientFactory.getClient(`${config.appIdUri}`);
    
    const requestUrl = `${config.hostUrl}/api/SearchSites?QueryText=${queryText}`;
    return this.client
      .get(requestUrl, AadHttpClient.configurations.v1)   
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then((items: any[]) => {
        const sites: Array<ISite> = new Array<ISite>();
        items.forEach(i => {
          sites.push({Id : i.id, Title : i.title, Url : i.url });
        });
        return sites;
      }); 
  }

  public async servicePrincipals(prefix: string): Promise<IApp[]> {
    this.client = await this.aadHttpClientFactory.getClient(`${config.appIdUri}`);
    const requestUrl = `http://localhost:7086/api/GetServicePrincipals?Prefix=${prefix}`;
    return this.client
      .get(requestUrl, AadHttpClient.configurations.v1)   
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then((items: any[]) => {
        const apps: Array<IApp> = new Array<IApp>();
        items.forEach(i => {
          apps.push({Id : i.id, DisplayName : i.displayName });
        });
        return apps;
      });
  }

  public async isSiteAdmin(userEMail: string, currentSiteId: string): Promise<boolean> {
    this.client = await this.aadHttpClientFactory.getClient(`${config.appIdUri}`);
    const requestUrl = `${config.hostUrl}/api/IsSiteAdmin?SiteId=${currentSiteId}&User=${userEMail}`;
    return this.client
      .get(requestUrl, AadHttpClient.configurations.v1)   
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then((items: boolean) => {        
        return items;
      }); 
  }

  public async grantPermissions(role: string, appId: string, site: ISite): Promise<any[]> {
    this.client = await this.aadHttpClientFactory.getClient(`${config.appIdUri}`);
    const requestUrl = `http://localhost:7086/api/ApplyPermimssion`;
    const requestBody = {      
      URL: site.Url,
      Permission: role,
      AppID: appId
    };
    return this.client
      .post(requestUrl, AadHttpClient.configurations.v1,
        { 
          body: JSON.stringify(requestBody) 
        }
      )   
      .then((response: HttpClientResponse) => {
        return response.json();
      }); 
  }
}