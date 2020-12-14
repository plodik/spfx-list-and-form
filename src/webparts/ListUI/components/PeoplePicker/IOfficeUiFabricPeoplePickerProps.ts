import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointUserPersona } from '../Models/OfficeUiFabricPeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IOfficeUiFabricPeoplePickerProps {
  context: WebPartContext;
  titleText: string;
  personSelectionLimit: number;
  onChange?: (items: SharePointUserPersona[]) => void;
  defaultSelectedUsers?: SharePointUserPersona[];
  disabled?: boolean;
}
