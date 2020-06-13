import { IListItem } from './IListItem';

export interface ISpfxReadWriteOperationsState {
  status: string;
  items: IListItem[];
  hideItemCreationConfirmationDialog: boolean;
}