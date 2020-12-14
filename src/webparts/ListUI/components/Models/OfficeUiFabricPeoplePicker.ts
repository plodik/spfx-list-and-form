import { IPersonaProps, IPersona } from "office-ui-fabric-react";

export interface IOfficeUiFabricPeoplePickerState {
    selectedUsers: SharePointUserPersona[];
}

export interface IEnsureUser {
    imageInitials: string;
    id: string;
    imageUrl: string;
    loginName: string;
    optionalText: string;
    secondaryText: string;
    tertiaryText: string;
    text: string;
}

export class SharePointUserPersona  implements IPersona {
    private _user:IEnsureUser;
    public get User(): IEnsureUser {
        return this._user;
    }

    public set User(user: IEnsureUser) {
        this._user = user;
        this.text = user.text;
        this.secondaryText = user.secondaryText;
        this.tertiaryText = user.tertiaryText;
        this.imageShouldFadeIn = true;
        this.imageUrl = user.imageUrl;
    }

    constructor (user: IEnsureUser) {
        this.User = user;
    }

    public id: number;
    public userName: string;
    public text: string;
    public secondaryText: string;
    public tertiaryText: string;
    public imageUrl: string;
    public imageShouldFadeIn: boolean;
}
