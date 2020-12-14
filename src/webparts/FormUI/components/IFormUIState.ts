import { SPListItem } from "@microsoft/sp-page-context";
import { CustomItem } from "./Models/ICustomItem";
import {IUser, User } from "../../../Models/User";

export interface IFormUIState {
  returnPageName?: string;

  actionIsNew?: boolean;
  actionIsEdit?: boolean;
  actionEditCustomItemID?: number;
  actionEditIsCustomItemLoadedOK?: boolean;
  actionEditIsCustomItemIDNumber?: boolean;
  actionEditCustomItemSPListItem?: SPListItem;
  actionEditIsCurrentUserAuthor?: boolean;

  formLoading?: boolean;
  showForm?: boolean;
  isFinalAndReadOnly?: boolean;

  actionFinishedMessage?: string;
  actionErrorMessage?: string;
  actionIsFinished?: boolean;

  customItem?: CustomItem;

  currentUser?: IUser;
  currentUserGroups?: any[];

  // stavy pro validatory
  isValid_DateCustomItem?: boolean;
  isValid_ItemType?: boolean;
  isValid_OurMarker?: boolean;

  // counters
  itemTypesAll?: any[];

  // test rest data
  testRestData?: any[];

  // approval process
  showApprovalActions?: boolean;
  showRejectedInfo?: boolean;
  showStartApproval?: boolean;
  showApprovalDetails?: boolean;
}
