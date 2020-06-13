import * as React from 'react';
import styles from './SpfxReadWriteOperations.module.scss';
import { ISpfxReadWriteOperationsProps } from './ISpfxReadWriteOperationsProps';
import { ISpfxReadWriteOperationsState } from './ISpfxReadWriteOperationsState';
import { IListItem } from './IListItem';
import {SPDataOperations} from '../../../DAL/SPDataOperations';
import {SPLogger} from '../../../Common/SPLogger';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { Dialog, DialogType, DialogFooter, IDialogContentProps } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton} from 'office-ui-fabric-react/lib/Button';

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
  color: 'white'
});

export default class SpfxReadWriteOperations extends React.Component<ISpfxReadWriteOperationsProps, ISpfxReadWriteOperationsState> {
  private oSPLogger: SPLogger;  
  private _columns: IColumn[];
  private titleForNewItem: string = '';
  private docIDForNewItem: string = '';
  private textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };  

  constructor(props: ISpfxReadWriteOperationsProps, state: ISpfxReadWriteOperationsState) {
    super(props);
    this.state = {
      status: this.listNotConfigured(this.props) ? 'Please configure list in Web Part properties' : 'Ready',
      items: [],
      hideItemCreationConfirmationDialog: true  
    };
    this._columns = [
      { key: 'ID', name: 'ID', fieldName: 'ID', minWidth: 50, maxWidth: 50, isResizable: true },
      { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 200, maxWidth: 500, isResizable: true },
      { key: 'DocID', name: 'DocID', fieldName: 'DocID', minWidth: 100, maxWidth: 200, isResizable: true }
    ];    
    this.oSPLogger = new SPLogger(this.props.context.serviceScope);
  }

  public componentWillReceiveProps(nextProps: ISpfxReadWriteOperationsProps): void {    
    this.setState({
      status: this.listNotConfigured(nextProps) ? 'Please configure list in Web Part properties' : 'Ready',
      items: []
    });
  }

  public render(): React.ReactElement<ISpfxReadWriteOperationsProps> {
 
    const disabled: string = this.listNotConfigured(this.props) ? styles.disabled : '';

    return (
      <div className={styles.spfxReadWriteOperations}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.subTitle}>
                SharePoint Deep Dive Assignment SPFX - 1
              </span>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <p></p>
              <span className={styles.subTitle}>
                Title:
              </span>
              <TextField
                className={exampleChildClass}                
                label=""
                onChange={this._onTitleChanged}
                styles={this.textFieldStyles}
                value={this.titleForNewItem}                
              />
              <p></p>
              <span className={styles.subTitle}>
                Doc ID:
              </span>
              <TextField 
                className={exampleChildClass}                               
                label=""
                onChange={this._onDocIDChanged}
                styles={this.textFieldStyles}
                value={this.docIDForNewItem}
              />
              <p></p>
              <a href="#" className={`${styles.button} ${disabled}`} onClick={() => this.createItem(this.titleForNewItem,this.docIDForNewItem)}>
                <span className={styles.label}>Create item</span>
              </a>
              <Dialog                
                hidden={this.state.hideItemCreationConfirmationDialog}
                onDismiss={this.toggleHideDialog}
                type= {DialogType.largeHeader}
                title= 'Item Added Successfully !!'
                closeButtonAriaLabel='Close'
                subText= {this.state.status}                                
              >
                <DialogFooter>
                  <PrimaryButton onClick={this.toggleHideDialog} text="OK" />                
                </DialogFooter>
              </Dialog>              
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <a href="#" className={`${styles.button} ${disabled}`} onClick={() => this.readItems()}>
                <span className={styles.label}>Read all items</span>
              </a>
            </div>
          </div>          
          <div className={styles.row}> 
                {this.state.status}
                 <p></p>             
              {this.state.items.length > 0 &&
                 <div>                 
                <DetailsList items={this.state.items}
                columns={this._columns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}/> 
                </div>               
              }            
          </div>
        </div>
      </div>
    );
  }

  private _onTitleChanged = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.titleForNewItem = text.trim();
  }

  private _onDocIDChanged = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.docIDForNewItem = text.trim();
  }

  private toggleHideDialog = () => {
    this.setState({      
      hideItemCreationConfirmationDialog: !this.state.hideItemCreationConfirmationDialog
    });    
  }


  private async createItem(title: string, docID: string){    
    try {
            this.setState({
              status: 'Creating item...',
              items: []
            });
            if(title.trim() === '') {
             title = `Item ${new Date()}`;
            }
            if(docID.trim() === '') {
             docID = 'DOCREP-10-9999';
            }
            let spDal = new SPDataOperations(this.props.context);            
            await spDal.createItem(this.props.listID,title,docID);
            this.setState({
              status: `Item with title ${title} successfully created`,
              items: [],
              hideItemCreationConfirmationDialog: false
            });
            this.titleForNewItem='';
            this.docIDForNewItem='';                     
        } catch (error) {    
            this.setState({
              status: 'Error while creating the item: ' + error,
              items: []
            });
            this.oSPLogger.logError(error);
        } 
  }
  
  private async readItems(){
    try {    
          let allItems: IListItem[];
          this.setState({
            status: 'Loading all items...',
            items: []
          });    
          let spDal = new SPDataOperations(this.props.context);            
          allItems = await spDal.loadListItems(this.props.listID);
          this.setState({
            status: `Successfully loaded ${allItems.length} items`,
            items: allItems
          });
        } catch (error) {    
          this.setState({
            status: 'Loading all items failed with error: ' + error,
            items: []
          });
          this.oSPLogger.logError(error);
  }
}  
  private listNotConfigured(props: ISpfxReadWriteOperationsProps): boolean {
    return props.listID === undefined ||
      props.listID === null ||
      props.listID.length === 0;
  }
}
