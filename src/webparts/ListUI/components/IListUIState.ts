import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IUser } from '../../../Models/User';

export interface IListUIState {
  items?: any[];
  columns?: IColumn[];
  status?: string;
  currentPage?: number;
  pageName?: string;
  pageSize?: number;
  order_ColumnName?: string;
  order_Direction?: string;

  dataApiCurrentLink?: string;
  dataApiNextLink?: string;
  pagingIsNextEnabled?: boolean;

  listDisplayType?: string;

  currentUser?: IUser;
  currentUserGroups?: any[];
  
  // filters for list
  statesAll?: any[];
  filterState?: any;
  filterOurMarker?: string;
}
