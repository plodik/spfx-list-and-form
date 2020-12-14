import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class UserDA {

  private context: IWebPartContext;

  constructor(context: IWebPartContext) {
    this.context = context;
  }

  public getUserIdByLoginName(loginName: string): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/siteusers?$filter=substringof(%27|${loginName}%27,LoginName)%20eq%20true&$select=id`, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
      })
        .then((response: SPHttpClientResponse): Promise<{ value: { Id: number }[] }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: { Id: number }[] }): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response.value[0].Id);
          }
        });
    });
  }

  public getUserGroupsByUserID(userID: number): Promise<any[]> {
    return new Promise<any[]>((resolve: (groups: any[]) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/GetUserById(${userID})/Groups`, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
      })
        .then((response: SPHttpClientResponse): Promise<{ value: { groupsData: any }[] }> => { return response.json(); }, (error: any): void => { reject(error); })
        .then((response: { value: { groupsData: any }[] }): void => { resolve(response.value); });
    });
  }

  public getCurrentUserDetails(): Promise<[string, string, string, string]> {
    // lists all available properties of user:
    // http://xxxx.sharepoint.com/sites/sampleListAndForm/_api/SP.UserProfiles.PeopleManager/GetMyProperties
    // filter to specific properties only:
    // http://xxxx.sharepoint.com/sites/sampleListAndForm/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=AccountName

    return new Promise<[string, string, string, string]>((resolve: (displayNameAndDepartmentAndEmail: [string, string, string, string]) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
      })
        .then((response: SPHttpClientResponse): Promise<{ value: any }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { value: any }): void => {
          var displayName: string;
          var email: string;
          var departmentNumber: string;
          var personalNumber: string;
          displayName = response["DisplayName"];
          email = response["Email"];

          var firstName: string = "";
          var lastName: string = "";

          try { personalNumber = response["AccountName"].split('\\')[1]; } catch (e) { personalNumber = "N/A"; }
          var properties: any[] = response["UserProfileProperties"];
          properties.forEach(property => {
            if (property.Key.toLowerCase() === "departmentnumber") {
              departmentNumber = this.StringNullableValueToString(property.Value);
            }
            if (property.Key.toLowerCase() === "firstname") {
              firstName = this.StringNullableValueToString(property.Value);
            }
            if (property.Key.toLowerCase() === "lastname") {
              lastName = this.StringNullableValueToString(property.Value);
            }
          });

          resolve([displayName, departmentNumber, email, personalNumber]);
        });
    });
  }

  private StringNullableValueToString(stringValue): string {
    if (stringValue === undefined || stringValue === null) { return ""; } else { return stringValue; }
  }
}
