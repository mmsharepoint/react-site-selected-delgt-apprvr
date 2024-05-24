import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class SPService {
  private _spHttpClient: SPHttpClient;

  public static readonly serviceKey: ServiceKey<SPService> =
    ServiceKey.create<SPService>('react-application-nav-sp', SPService);

  constructor(serviceScope: ServiceScope) {  
    serviceScope.whenFinished(async () => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
    });
  }

  public async isSiteAdmin(userEMail: string ,currentSiteUrl: string): Promise<boolean> {
    // const requestUrl = currentSiteUrl + `/_api/web/lists/GetByTitle('User Information List')/items?$select=Name,EMail,IsSiteAdmin&$filter=IsSiteAdmin eq 1`; //
    const requestUrl = currentSiteUrl + `/_api/web/lists/GetByTitle('User Information List')/items?$select=Name,EMail,IsSiteAdmin&$filter=EMail eq '${userEMail}'`;
    return this._spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((jsonResponse: any): any => {
        console.log(jsonResponse.value[0].IsSiteAdmin);
        return jsonResponse.value[0].IsSiteAdmin;
      });
  }
}