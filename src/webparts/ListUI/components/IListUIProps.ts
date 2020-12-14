import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IItemProp } from '../ListUIWebPart';

export interface IListUIProps {
  context: WebPartContext;
  currentUserLoginName: string;
  siteUrl: string;
  formPageName: string;
  listName: string;
  selectedColumns: IItemProp[];
  needsConfiguration: boolean;
  configureWebPart: () => void;
  displayMode: DisplayMode;
  pageSize: number;
  listDisplayType: string;
}
