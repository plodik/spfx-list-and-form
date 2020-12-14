var __cdnBasePath =  require('../../../../config/write-manifests.json');

export default class Constants {
  public static readonly _ListName_evidence: string = "CustomList";
  public static readonly _ListName_itemType: string = "ItemType";

  public static readonly _mainGarantGroupLoginName: string = "MainGarants";
  public static readonly _limitedGarantGroupLoginName: string = "LimitedGarants";

  public static readonly _CDNBASEPATH: string = __cdnBasePath["cdnBasePath"];

  public static readonly State_New: number = 1;
  public static readonly State_ForApproval: number = 2;
  public static readonly State_Rejected: number = 3;
  public static readonly State_Accepted: number = 4;
  public static readonly State_Revoked: number = 5;
}
