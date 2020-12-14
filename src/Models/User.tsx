export interface IUser {
  ID: number;
  LoginName: string;
  DepartmentNumber: string;
  DisplayName: string;
  Email: string;
  PersonalNumber: string;
  isMainGarant: boolean;
  isLimitedGarant: boolean;
}

export class User implements IUser {
  ID: number;
  LoginName: string;
  DepartmentNumber: string;
  DisplayName: string;
  Email: string;
  PersonalNumber: string;
  isMainGarant: boolean;
  isLimitedGarant: boolean;
}
