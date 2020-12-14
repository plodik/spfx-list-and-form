import * as React from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import { IOfficeUiFabricPeoplePickerProps } from './IOfficeUiFabricPeoplePickerProps';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IOfficeUiFabricPeoplePickerState, SharePointUserPersona, IEnsureUser } from '../Models/OfficeUiFabricPeoplePicker';
import { ISiteUserInfo, IWebEnsureUserResult } from '@pnp/sp/site-users/types';
export default class OfficeUiFabricPeoplePicker extends React.Component<IOfficeUiFabricPeoplePickerProps, IOfficeUiFabricPeoplePickerState> {

  constructor(props: IOfficeUiFabricPeoplePickerProps, state: IOfficeUiFabricPeoplePickerState) {
    super(props);
    sp.setup({ spfxContext: this.props.context });
    this.state = { selectedUsers: [] };
  }

  public render(): React.ReactElement<IOfficeUiFabricPeoplePickerProps> {
    return (
      <div>
        <PeoplePicker
          context={this.props.context}
          titleText={this.props.titleText}
          personSelectionLimit={this.props.personSelectionLimit}
          // groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
          showtooltip={true}
          disabled={this.props.disabled}
          onChange={this._getPeoplePickerItems.bind(this)}
          defaultSelectedUsers={this.props.defaultSelectedUsers && this.props.defaultSelectedUsers.map((item: SharePointUserPersona) => { return item.userName; })}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000} />
      </div>
    );
  }

  private _getPeoplePickerItems(items: any[]) {
    let persons: SharePointUserPersona[] = items.map(p => new SharePointUserPersona(p as IEnsureUser));

    if (persons.length > 0) {
      persons.forEach(user => {
        sp.web.ensureUser(user.User.loginName)
          .then((ensureUserResult: IWebEnsureUserResult): void => {
            sp.web.siteUsers.getByLoginName(user.User.loginName).select("id", "UserPrincipalName").get()
              .then((userInfo: ISiteUserInfo): void => {
                user.id = userInfo.Id;
                user.userName = userInfo.UserPrincipalName;
                this.setState({ selectedUsers: persons })
                if (this.props.onChange) { this.props.onChange(persons); }
              });
          });
      });
    }
    else {
      this.setState({ selectedUsers: persons })
      if (this.props.onChange) { this.props.onChange(persons); }
    }
  }
}
