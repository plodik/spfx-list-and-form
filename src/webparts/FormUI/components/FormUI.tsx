import * as React from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/fields/types";
import "@pnp/sp/profiles";
import { ISiteUserInfo } from '@pnp/sp/site-users/types';
import styles from './FormUI.module.scss';
import { IFormUIState } from './IFormUIState';
import { IFormUIProps } from './IFormUIProps';
import { Dropdown, IDropdownOption, TextField, Checkbox, Label } from 'office-ui-fabric-react';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import OfficeUiFabricPeoplePicker from '../../ListUI/components/PeoplePicker/OfficeUiFabricPeoplePicker';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPListItem } from '@microsoft/sp-page-context';
import { jsPDF } from "jspdf";
import { SharePointUserPersona } from '../../ListUI/components/Models/OfficeUiFabricPeoplePicker';
import Constants from './Constants';
import { IUser, User } from '../../../Models/User'
import { CustomItem } from './Models/ICustomItem';
import { ItemType } from './Models/IItemType';
import UserDA from '../../../Utilities/UserDA';
import CustomItemDA from './DataAccess/CustomItemDA';
import ColumnNames_CustomItem from './DataAccess/ColumnNames_CustomItem';
import ColumnNames_ItemType from './DataAccess/ColumnNames_ItemType';

export default class FormUI extends React.Component<IFormUIProps, IFormUIState> {

  private customItemDA: CustomItemDA;
  private userDA: UserDA;

  constructor(props: IFormUIProps) {
    super(props);
    sp.setup({ spfxContext: this.props.context });

    //var fullUrl = window.location.href;
    //var querystringUrl = window.location.search;

    let urlParams = new URLSearchParams(window.location.search.toLowerCase());

    let bActionIsEdit: boolean = false;
    let bActionIsNew: boolean = false;
    let numPOID: number = NaN;
    let bActionEditIsCustomItemIDNumber: boolean = false;
    let bActionEditCustomItemSPListItem: SPListItem = null;
    let bCustomItem: CustomItem = null;

    this.customItemDA = new CustomItemDA(this.props.context);
    this.userDA = new UserDA(this.props.context);

    // default action is NEW. If a=edit in querystring, then EDIT
    bActionIsNew = true;
    if (urlParams.has('a')) {
      if (urlParams.get('a') === "edit") {
        // EDIT action
        bActionIsEdit = true;
        bActionIsNew = false;
        // find POID
        if (urlParams.has('poid')) {
          numPOID = parseInt(urlParams.get('poid'));
          if (isNaN(numPOID)) { bActionEditIsCustomItemIDNumber = false; }
          else { bActionEditIsCustomItemIDNumber = true; }
        }
        else { bActionEditIsCustomItemIDNumber = false; }
      }
      else { bCustomItem = new CustomItem(); }
    }
    else { bCustomItem = new CustomItem(); }

    // return page url can be in 'r' attribute of querystring
    var aReturnPageName: string;
    if (urlParams.has('r')) { aReturnPageName = urlParams.get('r'); }
    if (aReturnPageName === undefined || aReturnPageName === null || aReturnPageName.length < 1) { aReturnPageName = this.props.listPageName.replace(".aspx", ""); } // if nothing found in querystring, fallback to default page name

    this.state = {
      showForm: false,
      formLoading: true,

      returnPageName: aReturnPageName,
      actionIsFinished: false,
      actionEditCustomItemSPListItem: bActionEditCustomItemSPListItem,
      actionEditIsCustomItemIDNumber: bActionEditIsCustomItemIDNumber,
      actionEditIsCustomItemLoadedOK: false,
      actionEditCustomItemID: numPOID,
      actionIsEdit: bActionIsEdit,
      actionIsNew: bActionIsNew,
      actionEditIsCurrentUserAuthor: false,

      customItem: bCustomItem,

      isFinalAndReadOnly: false,
      showApprovalActions: false,
      showRejectedInfo: false,
      showStartApproval: false,
      showApprovalDetails: false,

      isValid_ItemType: false,
      isValid_DateCustomItem: false,
      isValid_OurMarker: false,

      itemTypesAll: [],
      testRestData: [],
    };
  }

  public componentWillMount(): void {
    // componentWillMount
  }

  public componentDidMount(): void {
    // componentDidMount
    var currentUser: IUser = new User();

    this.userDA.getUserIdByLoginName(this.props.currentUserLoginName)
      .then((xCurrentUserID: number): Promise<[string, string, string, string]> => {
        currentUser.ID = xCurrentUserID;
        return this.userDA.getCurrentUserDetails();
      })
      .then((xDisplayNameAndDepartmentNumberAndEmailAndPersonalNumber: [string, string, string, string]): Promise<any[]> => {
        currentUser.DisplayName = xDisplayNameAndDepartmentNumberAndEmailAndPersonalNumber[0];
        currentUser.DepartmentNumber = xDisplayNameAndDepartmentNumberAndEmailAndPersonalNumber[1];
        currentUser.Email = xDisplayNameAndDepartmentNumberAndEmailAndPersonalNumber[2];
        currentUser.PersonalNumber = xDisplayNameAndDepartmentNumberAndEmailAndPersonalNumber[3];
        return this.userDA.getUserGroupsByUserID(currentUser.ID);
      })
      .then((groups: any[]): Promise<boolean> => {
        /* check if is member of MainGarant group */
        var isMainGarant: boolean = false;
        var isLimitedGarant: boolean = false;
        groups.forEach(group => {
          if (group.LoginName == Constants._mainGarantGroupLoginName) { isMainGarant = true; }
          if (group.LoginName == Constants._limitedGarantGroupLoginName) { isLimitedGarant = true; }
        });
        currentUser.isMainGarant = isMainGarant;
        currentUser.isLimitedGarant = isLimitedGarant;
        this.setState({ currentUser: currentUser, currentUserGroups: groups });
        return new Promise<boolean>(resolve => { resolve(true) });
      })
      .then((success: boolean): Promise<boolean> => {
        // read counters - item types
        return this.readCounters();
      })
      .then((success: boolean): void => {
        // read SPListItem from CustomItem for EDIT action
        if (this.state.actionIsEdit && this.state.actionEditIsCustomItemIDNumber) {
          this.GetCustomItemItemByID(this.state.actionEditCustomItemID);
        }
        // default values for NEW CustomItem
        if (this.state.actionIsNew) {
          let aCustomItem = this.state.customItem;
          aCustomItem.StateId = Constants.State_New;
          aCustomItem.IsApproved = false;
          aCustomItem.IsForApproval = false;
          aCustomItem.IsRejected = false;
          aCustomItem.IsRevoked = false;
          this.setState({
            formLoading: false,
            showForm: true,
            customItem: aCustomItem,
          });
        }
      }).catch((error: any): void => {
        this.setState({
          formLoading: false,
          showForm: false,
          actionErrorMessage: 'Form Initialization failed, contact administrator: ' + error
        });
      });
  }

  public componentWillReceiveProps(nextProps: IFormUIProps): void {
    // empty, not needed now
  }

  public render() {

    let { formLoading, showStartApproval, currentUser, showApprovalActions, showApprovalDetails, showRejectedInfo, customItem, itemTypesAll, testRestData, actionEditIsCurrentUserAuthor, actionErrorMessage, actionFinishedMessage, actionIsFinished, actionEditCustomItemID, actionIsEdit, actionIsNew, actionEditIsCustomItemLoadedOK, actionEditIsCustomItemIDNumber, showForm } = this.state;

    return (
      <div className={styles.FormUI}>
        <div>
          <div>
            <div>
              <div hidden={!actionIsFinished}>
                {actionFinishedMessage}
              </div>
              <div hidden={actionIsFinished}>
                <div className={styles.status}>
                  {actionErrorMessage}
                </div>
                <div className={styles.formHeader}>
                  {(actionIsNew === true) &&
                    <h1>New Custom Item</h1>
                  }
                  {(actionIsEdit === true) &&
                    <h1>Edit of Custom Item ID {actionEditCustomItemID.toString()}</h1>
                  }
                </div>
                {(formLoading === true) &&
                  <div className={styles.formLoading}>
                    <div className={styles.formLoading}>Load form, please wait...</div>
                  </div>
                }
                {(actionIsEdit === true && actionEditIsCustomItemIDNumber === false) &&
                  <div className={styles.alert}>
                    <div className={styles.alert}>ID is not number</div>
                  </div>
                }
                {(actionIsEdit === true && formLoading === false && actionEditIsCustomItemLoadedOK === false) &&
                  <div className={styles.alert}>
                    <div className={styles.alert}>Custom Item not found</div>
                  </div>
                }
                {(showForm === true && formLoading === false) &&
                  <div>
                    {(actionIsEdit && actionEditIsCurrentUserAuthor === false) &&
                      <div className={styles.YouAreNotAuthorDiv}>
                        You are not author of this item. No edit possible.
                      </div>
                    }
                    {(actionIsEdit && actionEditIsCurrentUserAuthor === true) &&
                      <div className={styles.YouAreAuthorDiv}>
                        You are author of this item. Edit is possible.
                      </div>
                    }
                    <div>
                      <strong>Your name is: </strong> {currentUser.DisplayName}
                    </div>
                    <br /><br />
                    <div>
                      TEST External REST API call:<br />
                      <PrimaryButton
                        onClick={this._loadTestRestData.bind(this)}
                        text="Call External REST API"></PrimaryButton>
                      <Dropdown
                        id="testRestData"
                        options={testRestData && testRestData.map((item: any) => { return { key: item.id, text: item.name }; })}
                        ariaLabel="Test data"
                        label="Test data"
                        hidden={testRestData && testRestData.length > 0 ? false : true}
                      ></Dropdown>
                    </div>
                    <br /><br />
                    <div className={styles.Date}>
                      <Label>Date</Label>
                      <DatePicker
                        firstDayOfWeek={DayOfWeek.Monday}
                        strings={DayPickerStrings}
                        value={customItem.DateCustomItem}
                        onSelectDate={this._DK_DatePickerOnChange_DateCustomItem.bind(this)}
                        formatDate={this._onFormatDate}
                        placeholder="choose date..."
                        ariaLabel="choose date"
                        disabled={this.state.isFinalAndReadOnly}
                      />
                      {!this.state.isValid_DateCustomItem &&
                        <div className={styles.errorMessage}>Date is required</div>
                      }
                    </div>

                    <div className={styles.Person_Manager}><OfficeUiFabricPeoplePicker
                      context={this.props.context}
                      onChange={this._Person_Manager_PeoplePickerOnChange.bind(this)}
                      defaultSelectedUsers={this.state.customItem.Person_Manager}
                      titleText="Person Manager"
                      personSelectionLimit={1}
                      disabled={this.state.isFinalAndReadOnly}>
                    </OfficeUiFabricPeoplePicker>
                      <div className={styles.Person_desc}>Responsible person manager</div>
                    </div>

                    <div className={styles.ItemType}>
                      <Dropdown
                        id="ItemType"
                        selectedKey={customItem.ItemTypeId}
                        onChanged={this._ItemTypeOnChanged.bind(this)}
                        options={itemTypesAll && itemTypesAll.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                        ariaLabel="Item type"
                        label="Item type"
                        disabled={this.state.isFinalAndReadOnly}
                      ></Dropdown>
                      {!this.state.isValid_ItemType &&
                        <div className={styles.errorMessage}>Item type is required</div>
                      }
                    </div>

                    <div className={styles.OurMarker}><TextField
                      value={customItem.OurMarker}
                      onChange={(newValue) => this._textBoxInputChanged('OurMarker', newValue)}
                      id="OurMarker"
                      ariaLabel="Our marker"
                      label="Our marker"
                      required={true}
                      placeholder="Our marker..."
                      disabled={this.state.isFinalAndReadOnly}
                    ></TextField>
                      {!this.state.isValid_OurMarker &&
                        <div className={styles.errorMessage}>Our marker is required</div>
                      }
                    </div>

                    <div className={styles.floatClearer}></div>
                    <br /><br />

                    {(actionIsNew === true) &&
                      <PrimaryButton
                        disabled={!this.isValid()}
                        onClick={this._actionAddNewOnClicked.bind(this)}
                        text="Add new"></PrimaryButton>
                    }
                    {(actionIsEdit === true && customItem.IsForApproval === false && customItem.IsApproved === false && actionEditIsCurrentUserAuthor) &&
                      <PrimaryButton
                        disabled={!this.isValid()}
                        onClick={this._actionEditSaveOnClicked.bind(this)}
                        text="Save changes"></PrimaryButton>
                    }
                    <PrimaryButton
                      onClick={this._actionCancelOnClicked.bind(this)}
                      text="Cancel"
                      className={styles.buttonCancel}></PrimaryButton>

                    <br /><br />

                    {(actionIsEdit === true) &&
                      <div className={styles.StateDiv}>
                        State: <strong>{customItem.StateTitle}</strong>
                      </div>
                    }

                    <br />

                    {(showStartApproval === true) &&
                      <div>
                        <div className={styles.rejectedInfoDiv}>
                          <Label>When you finish the item changes, you can send it for approval.</Label>
                          <PrimaryButton
                            disabled={!this.isValid()}
                            onClick={this._actionAskForApprovalOnClicked.bind(this)}
                            text="Ask for approval"></PrimaryButton>
                        </div>
                      </div>
                    }

                    {(showRejectedInfo === true) &&
                      <div>
                        <div className={styles.rejectedInfoDiv}>
                          <div>
                            Rejection reason: <strong>{customItem.RejectionReason}</strong>
                          </div>
                          <PrimaryButton
                            onClick={this._actionAskForApprovalOnClicked.bind(this)}
                            text="Ask for approval again"></PrimaryButton>
                        </div>
                      </div>
                    }

                    {(showApprovalActions === true) &&
                      <div>
                        <div className={styles.rejectDiv}>
                          <Label>Sending it back to author for edit.</Label>
                          <TextField
                            value={customItem.RejectionReason}
                            onChange={(newValue) => this._textBoxInputChanged('RejectionReason', newValue)}
                            id="RejectionReason"
                            ariaLabel="Rejection reason"
                            label="Rejection reason"
                          ></TextField>
                          <PrimaryButton
                            onClick={this._actionRejectOnClicked.bind(this)}
                            text="Reject for edit"></PrimaryButton>
                        </div>
                        <div className={styles.approveDiv}>
                          <Label>Approve</Label>
                          <PrimaryButton
                            onClick={this._actionApproveOnClicked.bind(this)}
                            text="Approve"></PrimaryButton>
                        </div>
                      </div>
                    }

                    {(showApprovalDetails === true) &&
                      <div className={styles.approvalDetailsDiv}>
                        <div>
                          Custom Item was approved at <strong>{customItem.DateApproval.toLocaleDateString()}</strong>
                        </div>
                        {(actionEditIsCurrentUserAuthor) &&
                          <div>
                            <PrimaryButton
                              onClick={this._actionGeneratePrintPDFOnClicked.bind(this)}
                              text="Prepare the printable version (PDF)"></PrimaryButton>
                          </div>
                        }
                      </div>
                    }

                    <br></br>

                    {(showApprovalDetails === true) &&
                      <div className={styles.approvedDiv}>
                        <div>
                          <strong>Approved, final...</strong>
                        </div>
                      </div>
                    }
                  </div>
                }
              </div>
            </div>
          </div>
        </div>
      </div >
    );
  }

  private _loadTestRestData(): void {
    this.props.context.spHttpClient.get("https://gorest.co.in/public-api/products", SPHttpClient.configurations.v1, { //headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' */}
    })
      .then((response: SPHttpClientResponse): Promise<{ value: any[] }> => { if (response.status === 404) { return null; } else { return response.json(); } })
      .then((items: any): void => { this.setState({ testRestData: items.data }); }, (error: any): void => { /*resolve(null);*/ });
  }

  private readCounters(): Promise<boolean> {
    // https://pnp.github.io/pnpjs/sp/items/
    return new Promise<boolean>((resolve: (result: boolean) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle(Constants._ListName_itemType).items.select("ID", "Title", ColumnNames_ItemType._TextItemType).orderBy("Title", true).get()
        .then(response => {
          this.setState({ itemTypesAll: response });
          resolve(true);
        })
        .catch((error: any): void => {
          this.setState({
            itemTypesAll: [],
            actionErrorMessage: 'Loading all "' + Constants._ListName_itemType + '" failed with error: ' + error
          });
        });
    });

    /*
    // version to read items from Choice column instead of Lookup list
    return new Promise<boolean>((resolve: (result: boolean) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle(Constants._ListName_evidence).fields.getByTitle("ColumnName").select('Choices').get()
        .then((response: any) => {
          this.setState({ itemTypesAll: response.Choices });
          resolve(true);
        })
        .catch((error: any): void => {
          this.setState({
            typyMajetkuAll: [],
            actionErrorMessage: 'Loading all "ColumnName" failed with error: ' + error
          });
        });
    });
    */
  }

  private isValid(): boolean {
    return (
      this.state.isValid_OurMarker &&
      this.state.isValid_DateCustomItem &&
      this.state.isValid_ItemType
    ); // todo other required fields
  }

  private _actionAddNewOnClicked() {
    var aCustomItem: CustomItem = this.state.customItem;
    this.customItemDA.CustomItem_CreateNew(this.state.customItem)
      .then((createdPoID: number): void => { this.RedirectOrThankYou("Custom Item was saved, ID: " + createdPoID); })
      .catch((error: any): void => { this.setState({ actionErrorMessage: `Error creating item: ${error}` }); });
  }

  private _actionEditSaveOnClicked() {
    var aCustomItem: CustomItem = this.state.customItem;
    this.customItemDA.CustomItem_Update(aCustomItem, this.state.actionEditCustomItemID)
      .then((updatedPoID: number): void => { this.RedirectOrThankYou("Custom Item was saved, ID: " + this.state.actionEditCustomItemID); })
      .catch((error: any): void => { this.setState({ actionErrorMessage: `Error updating item: ${error}` }); });
  }

  public RedirectOrThankYou(thankYouMessage: string) {
    const url = `${this.props.siteUrl}/SitePages/${this.props.listPageName}?state=1`;
    FormUI.redirect(url, false);
    // if we like to show ThankYou message instead of redirect:
    //this.setState({ actionIsFinished: true, actionFinishedMessage: thankYouMessage }, function () { });
  }

  // workaround method for redirect
  public static redirect(url: string, newTab?: boolean) {
    // Create a hyperlink element to redirect so that SharePoint uses modern redirection
    const link = document.createElement('a');
    link.href = url;
    link.className = styles.hidden;
    link.target = newTab ? '_blank' : '';
    document.body.appendChild(link);
    link.click();
  }

  private _actionCancelOnClicked() {
    const url = `${this.props.siteUrl}/SitePages/${this.state.returnPageName}.aspx`;
    FormUI.redirect(url, false);
  }

  private _actionAskForApprovalOnClicked() {
    this.customItemDA.CustomItem_Update(this.state.customItem, this.state.actionEditCustomItemID)
      .then((updatedPoID: number): Promise<number> => { return this.customItemDA.CustomItem_AskForApproval(this.state.actionEditCustomItemID); })
      .then((updatedPoID: number): void => { this.RedirectOrThankYou("Custom Item was sent to approval, ID: " + this.state.actionEditCustomItemID); })
      .catch((error: any): void => { this.setState({ actionErrorMessage: `Error AskForApproval item: ${error}` }); });
  }

  private _actionRevokeOnClicked() {
    this.customItemDA.CustomItem_Revoke(this.state.actionEditCustomItemID)
      .then((updatedPoID: number): void => { this.RedirectOrThankYou("Custom Item was revoked, ID: " + this.state.actionEditCustomItemID); })
      .catch((error: any): void => { this.setState({ actionErrorMessage: `Error Revoke item: ${error}` }); });
  }

  public getFormUIURL(CustomItemID: number): string {
    var fullUrl = window.location.href;
    var querystringUrl = window.location.search;
    return fullUrl.replace(querystringUrl, "") + "?A=Edit&PoID=" + CustomItemID.toString();
  }

  private _actionGeneratePrintPDFOnClicked() {
    let aCustomItem = this.state.customItem;
    var aItemType: ItemType;
    this.GetItemTypeItemByID(aCustomItem.ItemTypeId)
      .then((itemType: ItemType): void => {
        aItemType = itemType;
        // 'a4'  : [ 595.28,  841.89] size in pt
        var doc = new jsPDF('p', 'pt', 'a4');
        doc.setProperties({ title: 'CustomItem' });
        // constants for positioning
        var lMargin = 50; // left margin in pt
        var rMargin = 50; // right margin in pt
        var pdfInMM = 595;  // width of A4 in pt
        var pageCenter = pdfInMM / 2;

        // start to build the document
        doc.setFont('Arial', 'bold');
        doc.setFontSize(16);

        doc.setLineWidth(1);
        doc.line(50, 125, 550, 125);

        doc.text('Datum:', 50, 190, null, { align: 'left' });
        doc.text(this.GetLocalStringCzech(aCustomItem.DateCustomItem), 150, 190, null, { align: 'left' });

        doc.setFontSize(13);
        doc.text('CustomItem ', pageCenter, 210, null, { align: 'left' });

        if (aCustomItem.Person_Manager !== undefined && aCustomItem.Person_Manager !== null && aCustomItem.Person_Manager.length > 0) {
          doc.setFontSize(11);
          doc.text('Manager:', 50, 250, null, { align: 'left' });
          doc.text(aCustomItem.Person_Manager[0].text, 180, 250, null, { align: 'left' });
        }

        doc.setFontSize(12);
        doc.text("certificate,", pageCenter, 360, null, { align: 'center' });

        doc.text(aItemType.TextItemType, pageCenter, 450, null, { align: 'center' });

        // footer
        doc.line(50, 800, 550, 800);
        doc.text("Phone: " + "123456", pageCenter, 815, null, { align: 'center' });

        // save doc for download
        doc.save('CustomItem_' + this.state.actionEditCustomItemID + '.pdf');
      });
  }

  private _actionApproveOnClicked() {
    this.customItemDA.CustomItem_Update(this.state.customItem, this.state.actionEditCustomItemID)
      .then((updatedPoID: number): Promise<number> => { return this.customItemDA.CustomItem_Approve(this.state.actionEditCustomItemID); })
      .then((updatedPoID: number): void => { this.RedirectOrThankYou("Custom Item was approved, ID: " + this.state.actionEditCustomItemID); });
  }

  private _actionRejectOnClicked() {
    this.customItemDA.CustomItem_Update(this.state.customItem, this.state.actionEditCustomItemID)
      .then((updatedPoID: number): Promise<number> => { return this.customItemDA.CustomItem_Reject(this.state.actionEditCustomItemID, this.state.customItem.RejectionReason); })
      .then((updatedPoID: number): void => { this.RedirectOrThankYou("Custom Item was rejected, ID: " + this.state.actionEditCustomItemID); });
  }

  private _DK_DatePickerOnChange_DateCustomItem(selectedDate: Date) { this._DK_DatePickerOnChange("DateCustomItem", selectedDate); }

  private _DK_DatePickerOnChange(columnName, selectedDate: Date) {
    let aCustomItem = this.state.customItem;
    var aIsValid_DateCustomItem: boolean = this.state.isValid_DateCustomItem;
    switch (columnName) {
      case "DateCustomItem": { aCustomItem.DateCustomItem = selectedDate; aIsValid_DateCustomItem = true; break; }
    };

    if (columnName === "DateCustomItem") { this.setState({ customItem: aCustomItem, isValid_DateCustomItem: aIsValid_DateCustomItem, }); }
    else {
      this.setState({
        customItem: aCustomItem,
        isValid_DateCustomItem: aIsValid_DateCustomItem,
      });
    }
  }

  public _textBoxInputChanged(columnName, newValue) {
    let newValueStr: string = newValue.target.value;
    let aCustomItem = this.state.customItem;
    var aIsValid_OurMarker: boolean = this.state.isValid_OurMarker;

    switch (columnName) {
      case "OurMarker": { aCustomItem.OurMarker = newValueStr; if (newValueStr.length > 0) { aIsValid_OurMarker = true; } else { aIsValid_OurMarker = false; }; break; }
      case "RejectionReason": { aCustomItem.RejectionReason = newValueStr; break; }
    };

    this.setState({
      customItem: aCustomItem,
      isValid_OurMarker: aIsValid_OurMarker,
    });
  }

  public GetSharePointUserPersonsFromUserID(userID: number): Promise<SharePointUserPersona[]> {
    return new Promise<SharePointUserPersona[]>((resolve: (users: SharePointUserPersona[]) => void, reject: (error: any) => void): void => {
      sp.web.siteUsers.getById(userID).get()
        .then((userInfo: ISiteUserInfo): string => { return userInfo.LoginName })
        .then((loginName: string): any => { return sp.profiles.getPropertiesFor(loginName); })
        .then((userProfile: any): void => {
          let uprops = {};
          userProfile.UserProfileProperties.map(function (val) { uprops[val.Key] = val.Value });
          var person: SharePointUserPersona[] = [];
          var user: any = ({
            id: userProfile.AccountName,
            imageInitials: "",
            imageUrl: uprops['PictureURL'],
            loginName: userProfile.AccountName,
            optionalText: "",
            secondaryText: uprops['UserName'],
            tertiaryText: "",
            text: userProfile.DisplayName
          });
          person[0] = new SharePointUserPersona(user);
          person[0].id = userID;
          person[0].userName = uprops['UserName'];
          resolve(person);
        });
    });
  }

  private _ItemTypeOnChanged = (option: IDropdownOption, index?: number): void => {
    let aCustomItem = this.state.customItem;
    aCustomItem.ItemTypeId = Number(option.key);
    this.setState({
      customItem: aCustomItem, isValid_ItemType: true
    }, function () { });
  }

  private GetItemTypeItemByID(itemTypeID: number): Promise<ItemType> {

    return new Promise<ItemType>((resolve: (itemType: ItemType) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle(Constants._ListName_itemType).items.select("ID", ColumnNames_ItemType._Title, ColumnNames_ItemType._TextItemType).getById(itemTypeID).get()
        .then((item: any): void => {
          var aItemType: ItemType = {
            ID: item.ID,
            Title: item.Title,
            TextItemType: item[ColumnNames_ItemType._TextItemType],
          };
          // return it back
          resolve(aItemType);
        }, (error: any): void => { resolve(null); });
    });
  }

  private GetCustomItemItemByID(CustomItemID: number): void {
    var selectQuery: string[] = [];
    var expandQuery: string[] = [];

    // sloupce do SELECT a EXPAND - ktere se nacitaji ze SPList
    selectQuery.push("ID");
    selectQuery.push(ColumnNames_CustomItem._Title);
    selectQuery.push(ColumnNames_CustomItem._OurMarker);
    selectQuery.push(ColumnNames_CustomItem._DateCustomItem);
    selectQuery.push(ColumnNames_CustomItem._IsApproved);
    selectQuery.push(ColumnNames_CustomItem._IsForApproval);
    selectQuery.push(ColumnNames_CustomItem._IsRejected);
    selectQuery.push(ColumnNames_CustomItem._IsRevoked);
    selectQuery.push(ColumnNames_CustomItem._RejectionReason);
    selectQuery.push(ColumnNames_CustomItem._DateApproval);
    expandQuery.push(ColumnNames_CustomItem._State);
    selectQuery.push(ColumnNames_CustomItem._State + "/ID");
    selectQuery.push(ColumnNames_CustomItem._State + "/Title");
    expandQuery.push(ColumnNames_CustomItem._ItemType);
    selectQuery.push(ColumnNames_CustomItem._ItemType + "/ID");
    selectQuery.push(ColumnNames_CustomItem._ItemType + "/Title");
    expandQuery.push("Author");
    selectQuery.push("Author/ID");
    expandQuery.push(ColumnNames_CustomItem._Person_Manager);
    selectQuery.push(ColumnNames_CustomItem._Person_Manager + "/ID");

    const selectColumns = selectQuery === null || selectQuery === undefined || selectQuery.length === 0 ? "" : '$select=' + selectQuery.join();
    const expandColumns = expandQuery === null || expandQuery === undefined || expandQuery.length === 0 ? "" : '&$expand=' + expandQuery.join();

    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${Constants._ListName_evidence}')/items(${CustomItemID})?` + selectColumns + expandColumns;

    this.props.context.spHttpClient.get(url,
      SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
    }).then((response: SPHttpClientResponse): Promise<{ value: any[] }> => {
      if (response.status === 404) {
        this.setState({
          formLoading: false,
          actionEditIsCustomItemLoadedOK: false,
          actionEditCustomItemSPListItem: null,
          customItem: null,
        });
      }
      else {
        return response.json();
      }
    }).then((item: any): void => {
      // if not found, exit
      if (item == undefined) { return; }

      var authorUserID = item.Author.ID;
      var currentUserID = this.state.currentUser.ID;
      // console.log('authorUserID: ' + authorUserID.toString() + ' -> currentUserID:' + currentUserID.toString());

      var aCustomItem: CustomItem = {
        ID: item.ID,
        Title: item.Title,
        OurMarker: item[ColumnNames_CustomItem._OurMarker],
        ItemTypeId: Number(item[ColumnNames_CustomItem._ItemType]["ID"].toString()),
        ItemTypeTitle: item[ColumnNames_CustomItem._ItemType]["Title"].toString(),
        DateCustomItem: new Date(item[ColumnNames_CustomItem._DateCustomItem]),
        StateId: Number(item[ColumnNames_CustomItem._State]["ID"].toString()),
        StateTitle: item[ColumnNames_CustomItem._State]["Title"].toString(),
        DateApproval: new Date(item[ColumnNames_CustomItem._DateApproval]),
        RejectionReason: item[ColumnNames_CustomItem._RejectionReason],
        IsForApproval: item[ColumnNames_CustomItem._IsForApproval],
        IsApproved: item[ColumnNames_CustomItem._IsApproved],
        IsRejected: item[ColumnNames_CustomItem._IsRejected],
        IsRevoked: item[ColumnNames_CustomItem._IsRevoked],

        // people columns will be loaded later
        Person_Manager: null,

        // for multiline column, it is important to remove HTML formatting
        // sample: this.MultilineColumnValueToBasicString(item["Note"]),

      };

      // get user id
      var person_ManagerID: number = (item[ColumnNames_CustomItem._Person_Manager] === undefined ? null : item[ColumnNames_CustomItem._Person_Manager].ID);
      if (person_ManagerID !== null) {
        this.GetSharePointUserPersonsFromUserID(person_ManagerID).then((users: SharePointUserPersona[]): void => {
          aCustomItem.Person_Manager = users;
          this.setState({ customItem: aCustomItem });
        });
      }

      // when the date is empty in Sharepoint, it will become 1/1/1970, it needs to be set to NULL
      if (aCustomItem.DateCustomItem.toLocaleDateString() === new Date(0).toLocaleDateString()) { aCustomItem.DateCustomItem = null; }

      // check the values and set validator states
      var aIsValid_OurMarker: boolean = true;
      var aIsValid_ItemType: boolean = true;
      var aIsValid_DateCustomItem: boolean = true;
      if (aCustomItem.OurMarker === null || aCustomItem.OurMarker.length === 0) { aIsValid_OurMarker = false; };
      if (aCustomItem.DateCustomItem === null || aCustomItem.DateCustomItem === undefined) { aIsValid_DateCustomItem = false; };
      if (aCustomItem.ItemTypeId === null || aCustomItem.ItemTypeId === 0) { aIsValid_ItemType = false; };

      /* START SECURITY! */
      var actionEditIsCurrentUserAuthor: boolean = false;
      if (currentUserID === authorUserID) { actionEditIsCurrentUserAuthor = true; }

      // hide all section for approval/reject
      var aShowApprovalActions: boolean = false;
      var aShowRejectedInfo: boolean = false;
      var aShowStartApproval: boolean = false;
      var aShowApprovalDetails: boolean = false;
      var aIsFinalAndReadOnly: boolean = false;

      // start approval or rejected info sections only for author
      if (actionEditIsCurrentUserAuthor) {
        if (aCustomItem.IsRejected) { aShowRejectedInfo = true; }
        if (aCustomItem.IsRejected === false && aCustomItem.IsForApproval === false && aCustomItem.IsApproved === false) {
          aShowStartApproval = true;
        }
      }

      // approve/reject action only - when CustomItem is IsForApproval
      if (actionEditIsCurrentUserAuthor && aCustomItem.IsForApproval) { aShowApprovalActions = true; }

      // this is for everyone - is aproved
      if (aCustomItem.IsApproved) {
        aShowApprovalActions = false;
        aShowRejectedInfo = false;
        aShowStartApproval = false;
        aShowApprovalDetails = true;
        aIsFinalAndReadOnly = true;
      }

      // user is not author - read-only!
      if (actionEditIsCurrentUserAuthor === false) { aIsFinalAndReadOnly = true; }
      /* END SECURITY! */

      this.setState({
        customItem: aCustomItem,
        isFinalAndReadOnly: aIsFinalAndReadOnly,
        showForm: true,
        formLoading: false,
        showApprovalActions: aShowApprovalActions,
        showRejectedInfo: aShowRejectedInfo,
        showApprovalDetails: aShowApprovalDetails,
        showStartApproval: aShowStartApproval,
        actionEditIsCustomItemLoadedOK: true,
        actionEditCustomItemSPListItem: item,
        actionEditIsCurrentUserAuthor: actionEditIsCurrentUserAuthor,

        isValid_OurMarker: aIsValid_OurMarker,
        isValid_DateCustomItem: aIsValid_DateCustomItem,
        isValid_ItemType: aIsValid_ItemType,
      });
    }).catch((error: any): void => {
      this.setState({
        formLoading: false,
        showForm: false,
        actionEditIsCustomItemLoadedOK: false,
        actionEditCustomItemSPListItem: null,
        customItem: null,
        actionErrorMessage: 'Failed to load CustomItem ' + CustomItemID.toString() + ', contact administrator: ' + error
      });
    });
  }

  private MultilineColumnValueToBasicString(stringValue): string {
    if (stringValue === undefined || stringValue === null) { return ""; }
    else {
      var reg1 = new RegExp("<div class=\"ExternalClass[0-9A-F]+\">[^<]*", ""); var reg2 = new RegExp("</div>$", ""); var reg3 = new RegExp("<p>*", "");
      var reg4 = new RegExp("</p>$", ""); var reg5 = new RegExp("<br>*", "g");
      return stringValue.replace(reg1, "").replace(reg2, "").replace(reg3, "").replace(reg4, "").replace(reg5, "\n");
    }
  }

  private StringNullableValueToString(stringValue): string {
    if (stringValue === undefined || stringValue === null) { return ""; } else { return stringValue; }
  }

  private GetLocalStringCzech(dateToCzechString: Date): string {
    // warning, getMonth() returns 0-11 for month part, need +1
    return dateToCzechString.getDate() + '.' + (dateToCzechString.getMonth() + 1) + '.' + dateToCzechString.getFullYear();
  }

  private _Person_Manager_PeoplePickerOnChange(items: any[]) {
    let aCustomItem = this.state.customItem;
    aCustomItem.Person_Manager = items;
    this.setState({ customItem: aCustomItem });
  }

  private _onFormatDate = (date: Date): string => {
    return date.getDate() + '.' + (date.getMonth() + 1) + '.' + (date.getFullYear());
  };
}

const DayPickerStrings: IDatePickerStrings = {
  months: ['Leden', 'Únor', 'Březen', 'Duben', 'Květen', 'Červen', 'Červenec', 'Srpen', 'Září', 'Říjen', 'Listopad', 'Prosinec'],
  shortMonths: ['Led', 'Únr', 'Bře', 'Dub', 'Kvě', 'Čer', 'Čec', 'Srp', 'Zář', 'Říj', 'Lis', 'Pro'],
  days: ['neděle', 'pondělí', 'úterý', 'středa', 'čtvrtek', 'pátek', 'sobota'],
  shortDays: ['N', 'P', 'U', 'S', 'Č', 'P', 'S'],
  goToToday: 'Přejít na dnešek',
  prevMonthAriaLabel: 'Přejít na předcházející měsíc',
  nextMonthAriaLabel: 'Přejít na další měsíc',
  prevYearAriaLabel: 'Přejít na předcházející rok',
  nextYearAriaLabel: 'Přejít na další rok',
  isRequiredErrorMessage: 'Pole je povinné.',
  invalidInputErrorMessage: 'Neplatný formát datumu.'
};