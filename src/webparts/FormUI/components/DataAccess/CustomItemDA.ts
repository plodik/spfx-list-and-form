import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { CustomItem } from '../Models/ICustomItem';
import Constants from '../Constants';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import ColumnNames_CustomItem from './ColumnNames_CustomItem';

export default class CustomItemDA {

  private context: IWebPartContext;
  constructor(context: IWebPartContext) { this.context = context; }

  public CustomItem_CreateNew(customItem: CustomItem): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      var body = {};
      body[ColumnNames_CustomItem._Title] = `CustomItem ${new Date().toLocaleString()}`;
      body[ColumnNames_CustomItem._State + "Id"] = Constants.State_New;
      body[ColumnNames_CustomItem._DateCustomItem] = customItem.DateCustomItem;
      body[ColumnNames_CustomItem._IsApproved] = false;
      body[ColumnNames_CustomItem._IsForApproval] = false;
      body[ColumnNames_CustomItem._IsRejected] = false;
      body[ColumnNames_CustomItem._IsRevoked] = false;
      body[ColumnNames_CustomItem._ItemType + "Id"] = customItem.ItemTypeId; // Id at the end, it is lookup column
      body[ColumnNames_CustomItem._OurMarker] = customItem.OurMarker;
      body[ColumnNames_CustomItem._RejectionReason] = customItem.RejectionReason;
      if (customItem.Person_Manager !== undefined && customItem.Person_Manager !== null && customItem.Person_Manager.length > 0) { body[ColumnNames_CustomItem._Person_Manager + "Id"] = customItem.Person_Manager[0].id; } // Id at the end, it is people column
      sp.web.lists.getByTitle(Constants._ListName_evidence).items.add(body)
        .then((response: any) => { resolve(response.Id); })
        .catch((error: any): void => { reject(`Error Create: ${error}`); });
    });
  }

  public CustomItem_Update(customItem: CustomItem, actionEditCustomItemID: number): Promise<number> {
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      var body = {};
      body[ColumnNames_CustomItem._Title] = customItem.Title;
      body[ColumnNames_CustomItem._DateCustomItem] = customItem.DateCustomItem;
      body[ColumnNames_CustomItem._ItemType + "Id"] = customItem.ItemTypeId; // Id at the end, it is lookup column
      body[ColumnNames_CustomItem._OurMarker] = customItem.OurMarker;
      body[ColumnNames_CustomItem._RejectionReason] = customItem.RejectionReason;
      if (customItem.Person_Manager !== undefined && customItem.Person_Manager !== null && customItem.Person_Manager.length > 0) { body[ColumnNames_CustomItem._Person_Manager + "Id"] = customItem.Person_Manager[0].id; } // Id at the end, it is people column
      sp.web.lists.getByTitle(Constants._ListName_evidence).items.getById(actionEditCustomItemID).update(body)
        .then((response: any) => { resolve(actionEditCustomItemID); })
        .catch((error: any): void => { reject(`Error Update: ${error}`); });
    });
  }

  public CustomItem_AskForApproval(actionEditCustomItemID: number): Promise<number> {
    return new Promise<number>((resolve: (result: number) => void, reject: (error: any) => void): void => {
      var body = {};
      body[ColumnNames_CustomItem._State + "Id"] = Constants.State_Accepted;
      body[ColumnNames_CustomItem._IsForApproval] = true;
      body[ColumnNames_CustomItem._IsRejected] = false;
      body[ColumnNames_CustomItem._IsRevoked] = false;
      sp.web.lists.getByTitle(Constants._ListName_evidence).items.getById(actionEditCustomItemID).update(body)
        .then((response: any) => { resolve(actionEditCustomItemID); })
        .catch((error: any): void => { reject(`Error AskForApproval: ${error}`); }); 
    });
  }

  public CustomItem_Approve(actionEditCustomItemID: number): Promise<number> {
    return new Promise<number>((resolve: (result: number) => void, reject: (error: any) => void): void => {
      var body = {};
      body[ColumnNames_CustomItem._State + "Id"] = Constants.State_Accepted;
      body[ColumnNames_CustomItem._IsApproved] = true;
      body[ColumnNames_CustomItem._DateApproval] = new Date();
      body[ColumnNames_CustomItem._IsForApproval] = false;
      body[ColumnNames_CustomItem._IsRejected] = false;
      body[ColumnNames_CustomItem._IsRevoked] = false;
      sp.web.lists.getByTitle(Constants._ListName_evidence).items.getById(actionEditCustomItemID).update(body)
        .then((response: any) => { resolve(actionEditCustomItemID); })
        .catch((error: any): void => { reject(`Error Approve: ${error}`); });
    });
  }

  public CustomItem_Reject(actionEditCustomItemID: number, RejectionReason: string): Promise<number> {
    return new Promise<number>((resolve: (result: number) => void, reject: (error: any) => void): void => {
      var body = {};
      body[ColumnNames_CustomItem._State + "Id"] = Constants.State_Rejected;
      body[ColumnNames_CustomItem._IsRejected] = true;
      body[ColumnNames_CustomItem._IsApproved] = false;
      body[ColumnNames_CustomItem._IsForApproval] = false;
      sp.web.lists.getByTitle(Constants._ListName_evidence).items.getById(actionEditCustomItemID).update(body)
        .then((response: any) => { resolve(actionEditCustomItemID); })
        .catch((error: any): void => { reject(`Error RejectForEdit: ${error}`); });
    });
  }

  public CustomItem_Revoke(actionEditCustomItemID: number): Promise<number> {
    return new Promise<number>((resolve: (result: number) => void, reject: (error: any) => void): void => {
      var body = {};
      body[ColumnNames_CustomItem._State + "Id"] = Constants.State_Revoked;
      body[ColumnNames_CustomItem._IsRevoked] = true;
      body[ColumnNames_CustomItem._IsApproved] = false;
      body[ColumnNames_CustomItem._IsForApproval] = false;
      sp.web.lists.getByTitle(Constants._ListName_evidence).items.getById(actionEditCustomItemID).update(body)
        .then((response: any) => { resolve(actionEditCustomItemID); })
        .catch((error: any): void => { reject(`Error Revoke: ${error}`); });
    });
  }
}