var __cdnBasePath = require('../../../../config/write-manifests.json');

export default class Constants {
  public static readonly _ListName_customList: string = "CustomList";
  public static readonly _ListName_customList_defaultOrderByColumnName: string = "DateCustomItem";

  public static readonly _ListName_state: string = "State";

  public static readonly _mainGarantGroupLoginName: string = "MainGarants";
  public static readonly _limitedGarantGroupLoginName: string = "LimitedGarants";

  public static readonly _CDNBASEPATH: string = __cdnBasePath["cdnBasePath"];
}
