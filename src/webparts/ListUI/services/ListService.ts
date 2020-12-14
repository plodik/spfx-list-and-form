import { IListService } from './IListService';
import { IList } from './IList';
import { IListColumn } from './IListColumn';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

export class ListService implements IListService {

    constructor(private context: IWebPartContext) {
    }

    public getColumns(listName: string): Promise<IListColumn[]> {
      var httpClientOptions : ISPHttpClientOptions = {};

      httpClientOptions.headers = {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
      };

      return new Promise<IListColumn[]>((resolve: (results: IListColumn[]) => void, reject: (error: any) => void): void => {

        // hides system read-only fields like ID
        // var apiURL = this.context.pageContext.web.serverRelativeUrl + `/_api/web/lists/GetByTitle('${listName}')/fields?$filter=TypeDisplayName ne 'Attachments' and Hidden eq false and ReadOnlyField eq false`;

        // shows all fields
        var apiURL = this.context.pageContext.web.serverRelativeUrl + `/_api/web/lists/GetByTitle('${listName}')/fields?$filter=TypeDisplayName ne 'Attachments' and Hidden eq false`;

        this.context.spHttpClient.get(apiURL,
          SPHttpClient.configurations.v1,
          httpClientOptions
          )
          .then((response: SPHttpClientResponse): Promise<{ value: IListColumn[] }> => {
            return response.json();
          })
          .then((listColumns: { value: IListColumn[] }): void => {
            resolve(listColumns.value);
          }, (error: any): void => {
            reject(error);
          });
      });
    }
}
