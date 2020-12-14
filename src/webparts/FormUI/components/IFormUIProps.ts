import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFormUIProps {
  context: WebPartContext;
  currentUserLoginName: string;
  siteUrl: string;
  listPageName: string;
}
