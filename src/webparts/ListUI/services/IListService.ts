import { IList } from './IList';
import { IListColumn } from './IListColumn'
export interface IListService {
  getColumns: (listName) => Promise<IListColumn[]>;
}
