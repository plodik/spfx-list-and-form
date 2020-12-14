import * as React from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import styles from './ListUI.module.scss';
import { IListUIState } from './IListUIState';
import { IListUIProps } from './IListUIProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DetailsList, IColumn, DetailsListLayoutMode as LayoutMode, ConstrainMode, CheckboxVisibility, SelectionMode, ColumnActionsMode } from 'office-ui-fabric-react/lib/DetailsList';
import { PrimaryButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, IDropdownOption, TextField } from 'office-ui-fabric-react';
import { Config } from './Config/Config';
import Paging from './Paging/Paging';
import Constants from './Constants';
import UserDA from '../../../Utilities/UserDA';
import { IUser, User } from '../../../Models/User';

export default class ListUI extends React.Component<IListUIProps, IListUIState> {
  private filterQuery: string[] = [];
  private selectQuery: string[] = [];
  private expandQuery: string[] = [];

  private userDA: UserDA;

  constructor(props: IListUIProps) {
    super(props);
    sp.setup({ spfxContext: this.props.context });

    this.userDA = new UserDA(this.props.context);

    var fullUrl = window.location.href;
    var siteUrl = this.props.context.pageContext.web.absoluteUrl
    var pageName = fullUrl.replace(siteUrl, '').replace('/SitePages/', '').replace('.aspx', '');

    let urlParams = new URLSearchParams(window.location.search.toLowerCase());

    var ourMarkerQS: string = '';
    if (urlParams.has('om')) { ourMarkerQS = urlParams.get('om'); };

    var stateQS: number = null;
    if (urlParams.has('state')) {
      switch (urlParams.get('state')) {
        case "1": { stateQS = 1; break; } // New
        case "2": { stateQS = 2; break; } // For approval
        case "3": { stateQS = 3; break; } // Rejected for edit
        case "4": { stateQS = 4; break; } // Accepted
        case "5": { stateQS = 5; break; } // Revoked
        default: { stateQS = null; break; }
      }
    };

    this.state = {
      items: [],
      columns: this.buildColumns(this.props),
      currentPage: 1,
      pageName: pageName,
      pageSize: this.props.pageSize,
      listDisplayType: this.props.listDisplayType,
      statesAll: [],
      filterOurMarker: ourMarkerQS,
      filterState: stateQS,
    };

    this._onPageUpdate = this._onPageUpdate.bind(this);
    this._onPageBackToFirst = this._onPageBackToFirst.bind(this);
    this._onPageNextPage = this._onPageNextPage.bind(this);
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
      .then((groups: any[]): Promise<number> => {
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
        return this.readFilterItems();
      })
      .then((statesCount: number): void => {
        const queryParam = this.buildQueryParams(this.props);
        const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items${queryParam}`;
        this.readItems(url);
      })
      .catch((error: any): void => {
        this.setState({ status: 'List Initialization failed, contact administrator: ' + error });
      });
  }

  public componentWillReceiveProps(nextProps: IListUIProps): void {
    this.setState({
      columns: this.buildColumns(nextProps),
      pageSize: nextProps.pageSize,
      currentPage: 1,
      dataApiCurrentLink: null,
      dataApiNextLink: null,
      pagingIsNextEnabled: false,
      listDisplayType: nextProps.listDisplayType,
      statesAll: []
    });

    const queryParam = this.buildQueryParams(nextProps);
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${nextProps.listName}')/items${queryParam}`;
    this.readItems(url);
  }

  public render() {

    const { needsConfiguration, configureWebPart } = this.props;
    let { items, columns, status, currentUser } = this.state;

    return (
      <div className={styles.ListUI}>
        <div>
          {needsConfiguration &&
            <Config configure={configureWebPart} {...this.props} />
          }
          {needsConfiguration === false &&
            <div>
              <div>
                <div>
                  {status !== "" &&
                    <div className={styles.status}>
                      {status}
                    </div>
                  }
                  {(status === undefined || status === null || status === "") &&
                    <div>
                      <div>
                        {this.props.listDisplayType !== "3" && // hide for 3 - only approved CustomItem
                          <Dropdown
                            selectedKey={this.state.filterState}
                            options={this.state.statesAll && this.state.statesAll.map((item: any) => { return { key: item.ID, text: item.Title }; })}
                            onChanged={this._FilterStateOnChanged.bind(this)}
                            id="filterState"
                            ariaLabel="State"
                            label="State"
                          ></Dropdown>
                        }
                        <TextField
                          value={this.state.filterOurMarker}
                          onChange={this._FilterOurMarkerOnChanged.bind(this)}
                          id="filterOurMarker"
                          ariaLabel="Our marker"
                          label="Our marker"
                        ></TextField>
                        <PrimaryButton
                          onClick={this._FilterClearOnClicked.bind(this)}
                          text="Clear filter"></PrimaryButton>
                      </div>
                      {items.length == 0 &&
                        <div className={styles.emptyResults}>No items...</div>
                      }
                      {items.length > 0 &&
                        <div>
                          <DetailsList
                            items={items}
                            columns={columns}
                            isHeaderVisible={true}
                            selectionMode={SelectionMode.single}
                            layoutMode={LayoutMode.justified}
                            constrainMode={ConstrainMode.unconstrained}
                            checkboxVisibility={CheckboxVisibility.hidden}
                            onColumnHeaderClick={this._onColumnClick.bind(this)}
                          />
                          <Paging
                            onBackToFirst={this._onPageBackToFirst}
                            onNextPage={this._onPageNextPage}
                            nextEnabled={this.state.pagingIsNextEnabled}
                            currentPage={this.state.currentPage} />
                        </div>
                      }
                    </div>
                  }
                </div>
              </div>
            </div>
          }
        </div>
      </div>
    );
  }

  private _FilterStateOnChanged = (option: IDropdownOption, index?: number): void => {
    this.setState({
      filterState: option.key
    }, function () { this._refreshList(); });
  }
  private _FilterOurMarkerOnChanged = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string): void => {
    this.setState({
      filterOurMarker: newValue
    }, function () { this._refreshList(); });
  }
  private _FilterClearOnClicked() {
    this.setState({
      filterState: null,
      filterOurMarker: ''
    }, function () { this._refreshList(); });
  }

  private readFilterItems(): Promise<number> {
    // https://pnp.github.io/pnpjs/sp/items/
    return new Promise<number>((resolve: (stavyCount: number) => void, reject: (error: any) => void): void => {
      sp.web.lists.getByTitle(Constants._ListName_state).items.select("ID", "Title").get()
        .then(response => {
          this.setState({ statesAll: response });
          resolve(response.length);
        })
        .catch((error: any): void => {
          this.setState({
            statesAll: [],
            status: 'Loading all "' + Constants._ListName_state + '" failed with error: ' + error
          });
        });
    });
  }

  private _refreshList() {
    const queryParam = this.buildQueryParams(this.props);
    const url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items${queryParam}`;
    this.readItems(url);
  }

  private readItems(url: string) {
    this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' } })
      .then((response: SPHttpClientResponse): Promise<{ value: any[] }> => { return response.json(); })
      .then((response: { value: any[] }): void => {
        var _nextLink = response['odata.nextLink'];
        var _currentLink = url;
        var _pagingIsNextEnabled: boolean = false;
        if (_nextLink === undefined || _nextLink === null) { _pagingIsNextEnabled = false; }
        else { _pagingIsNextEnabled = true; }
        this.setState({
          dataApiCurrentLink: _currentLink,
          dataApiNextLink: _nextLink,
          pagingIsNextEnabled: _pagingIsNextEnabled,
          items: response.value
        });
      }, (error: any): void => { this.setState({ items: [], status: 'Loading all items failed with error: ' + error }); });
  }

  private _onPageBackToFirst(): void { this._onPageUpdate(1); }
  private _onPageNextPage(pageNumber: number): void { this._onPageUpdate(pageNumber); }
  private _onPageUpdate(pageNumber: number) {
    this.setState({ currentPage: pageNumber });
    var url = this.state.dataApiCurrentLink;
    if (this.state.currentPage < pageNumber) {
      // move next
      url = this.state.dataApiNextLink;
    }
    else {
      // move first
      const queryParam = this.buildQueryParams(this.props);
      url = `${this.props.siteUrl}/_api/web/lists/GetByTitle('${this.props.listName}')/items${queryParam}`;
    }
    this.readItems(url);
  }

  private _onColumnClick(event: React.MouseEvent<HTMLElement>, column: IColumn) {
    let { order_ColumnName, order_Direction, columns } = this.state;

    var sortDir: string = "asc";
    if (order_ColumnName == column.key && order_Direction == "asc") { sortDir = "desc" }

    this.setState({
      order_ColumnName: column.key,
      order_Direction: sortDir,
      columns: columns!.map(col => {
        col.isSorted = (col.key === column.key);
        if (col.isSorted) { col.isSortedDescending = (sortDir === "desc"); }
        return col;
      })
    }, function () { this._refreshList(); });
  }

  private buildQueryParams(props: IListUIProps): string {
    this.filterQuery = [];
    this.selectQuery = [];
    this.expandQuery = [];
    props.selectedColumns.forEach(element => {
      var elementKey: string = element.key;
      if (elementKey.charAt(0) == "_") { elementKey = "OData_" + elementKey; }
      if (element.text === "Person or Group" || element.text === "Osoba nebo skupina" || element.text === "Lookup" || element.text === "Vyhledávání") {
        this.selectQuery.push(elementKey + "/Title");
        this.selectQuery.push(elementKey + "/ID");
        this.expandQuery.push(elementKey);
      }
      else {
        this.selectQuery.push(elementKey);
      }
    });

    // columns State, OurMarker and ID needs to be pulled from SP everytime even if then are not displayed!
    if (this.selectQuery.indexOf("ID") == -1) { this.selectQuery.push("ID") };
    if (this.selectQuery.indexOf("OurMarker") == -1) { this.selectQuery.push("OurMarker") };
    if (this.selectQuery.indexOf("State/Title") == -1) { this.selectQuery.push("State/Title") };
    if (this.selectQuery.indexOf("State/ID") == -1) { this.selectQuery.push("State/ID") };
    if (this.expandQuery.indexOf("State") == -1) { this.expandQuery.push("State") };

    // filter by Stav or OurMarker
    if (!(this.state.filterState === undefined || this.state.filterState === null)) {
      this.filterQuery.push("(StateId eq " + this.state.filterState + ")")
    }
    if (!(this.state.filterOurMarker === undefined || this.state.filterOurMarker === null)) {
      if (this.state.filterOurMarker.length > 0) {
        this.filterQuery.push("(startswith(OurMarker,'" + this.state.filterOurMarker + "'))")
      }
    }

    // listDisplayType:
    // !!! needs to be synced with PropertyPaneDropdown list in src\webparts\ListUI\ListUIWebPart.ts !!!
    // {key: '1', text: 'All items'},
    // {key: '2', text: 'Current user is author'},
    // {key: '4', text: 'State - Approved'},
    if (props.listDisplayType === "2") {
      this.filterQuery.push("(AuthorId eq " + this.state.currentUser.ID + ")")
    }
    if (props.listDisplayType === "3") {
      this.filterQuery.push("(StateId eq 4)") // hardcoded value - state = 4 - approved
    }

    const orderColumn = this.state.order_ColumnName === null || this.state.order_ColumnName === undefined || this.state.order_ColumnName.length === 0 ? '&$orderby=' + Constants._ListName_customList_defaultOrderByColumnName + ' desc' : '&$orderby=' + this.state.order_ColumnName + ' ' + this.state.order_Direction;

    const queryParam = `?$top=${props.pageSize}`;
    const filterColumns = this.filterQuery === null || this.filterQuery === undefined || this.filterQuery.length === 0 ? "" : '&$filter=' + this.filterQuery.join(' and ');
    const selectColumns = this.selectQuery === null || this.selectQuery === undefined || this.selectQuery.length === 0 ? "" : '&$select=' + this.selectQuery.join();
    const expandColumns = this.expandQuery === null || this.expandQuery === undefined || this.expandQuery.length === 0 ? "" : '&$expand=' + this.expandQuery.join();
    return queryParam + filterColumns + selectColumns + expandColumns + orderColumn;
  }

  private buildColumns(props: IListUIProps): IColumn[] {
    const columns: IColumn[] = [];

    const viewLinkColumn: IColumn = {
      key: 'Detail', name: 'Detail', fieldName: 'Detail',
      columnActionsMode: ColumnActionsMode.disabled,
      minWidth: 50, maxWidth: 50, isResizable: false,
      onRender: (rowitem) => {
        const url = `${this.props.siteUrl}/SitePages/${this.props.formPageName}?A=Edit&PmID=${rowitem.ID}&R=${this.state.pageName}`;
        return (
          <div>
            <IconButton className={styles.iconButtonDetail} iconProps={{ iconName: 'Trackers' }} title="Detail" ariaLabel="Detail" onClick={event => window.location.href = url} />
          </div>
        );
      }
    };
    columns.push(viewLinkColumn);

    props.selectedColumns.forEach(element => {
      var elementKey: string = element.key;
      if (elementKey.charAt(0) == "_") { elementKey = "OData_" + elementKey; }
      if (element.text.toLowerCase() === "person or group" || element.text.toLowerCase() === "osoba nebo skupina" || element.text.toLowerCase() === "lookup" || element.text.toLowerCase() === "vyhledávání") {
        const column: IColumn = {
          key: elementKey, name: element.displayName, fieldName: elementKey,
          minWidth: 100, maxWidth: 350, isResizable: true, data: 'string',
          onRender: (item: any) => {
            return (
              <span>
                {(item[elementKey] === undefined || item[elementKey] === null) ? "" : item[elementKey]["Title"]}
              </span>
            );
          }
        };
        columns.push(column);
      }
      else if (element.text.toLowerCase() === "date and time" || element.text.toLowerCase() === "datum a čas") {
        const column: IColumn = {
          key: elementKey, name: element.displayName, fieldName: elementKey,
          minWidth: 100, maxWidth: 350, isResizable: true, data: 'string',
          onRender: (item: any) => {
            return (
              <span>
                {(item[elementKey] === undefined || item[elementKey] === null) ? "" : new Date(item[elementKey]).toLocaleDateString()}
              </span>
            );
          }
        };
        columns.push(column);
      }
      else if (element.text.toLowerCase() === "currency" || element.text.toLowerCase() === "měna") {
        const column: IColumn = {
          key: elementKey, name: element.displayName, fieldName: elementKey,
          minWidth: 100, maxWidth: 350, isResizable: true, data: 'string',
          onRender: (item: any) => {
            return (
              <span>
                {item[elementKey] === null ? "" : item[elementKey] + " $"}
              </span>
            );
          }
        };
        columns.push(column);
      }
      else if (element.text.toLowerCase() === "yes/no" || element.text.toLowerCase() === "ano/ne") {
        const column: IColumn = {
          key: elementKey, name: element.displayName, fieldName: elementKey,
          minWidth: 100, maxWidth: 350, isResizable: true, data: 'string',
          onRender: (item: any) => {
            return (
              <span>
                {item[elementKey] === null || item[elementKey] === false ? "No" : "Yes"}
              </span>
            );
          }
        };
        columns.push(column);
      }
      else {
        const column: IColumn = {
          key: elementKey, name: element.displayName, fieldName: elementKey,
          minWidth: 100, maxWidth: 350, isResizable: true, data: 'string',
          isMultiline: element.text.toLowerCase() === "multiple lines of text" ? true : false
        };
        columns.push(column);
      }
    });
    return columns;
  }
}