import {sp} from '@pnp/sp/presets/all';
import {WebPartContext } from '@microsoft/sp-webpart-base';
import {SPLogger } from '../Common/SPLogger';
import {IListItem} from '../webparts/spfxReadWriteOperations/components/IListItem';

export class SPDataOperations {
  public context: WebPartContext;
  private oSPLogger: SPLogger;
  constructor(spContext: WebPartContext) {
    this.context = spContext; 
    sp.setup({
      spfxContext: this.context
    });
    this.oSPLogger = new SPLogger(this.context.serviceScope);
  }

  /**
   * Load all Items in the list
   *
   * @param listID The list selected in web part property
   */
  public async loadListItems(listID: string): Promise<IListItem[]> {
    let allItems: IListItem[];
    try {
        allItems = await sp.web.lists.getById(listID).items.select('Title', 'ID', 'DocID').get();
    } catch (error) {
      this.oSPLogger.logError(error);
    }
    return allItems;
  }

  /**
   * Add Item in the list
   *
   * @param listID The list selected in web part property
   */
  public async createItem(listID: string,title: string, docID: string){       
    try {
        await sp.web.lists.getById(listID).items.add({
            'Title': title,
            'DocID': docID
          });
    } catch (error) {
      this.oSPLogger.logError(error);
    }     
  }  
}