import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxReadWriteOperationsWebPartStrings';
import SpfxReadWriteOperations from './components/SpfxReadWriteOperations';
import { ISpfxReadWriteOperationsProps } from './components/ISpfxReadWriteOperationsProps';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface ISpfxReadWriteOperationsWebPartProps {
  listID: string;  
}

export default class SpfxReadWriteOperationsWebPart extends BaseClientSideWebPart <ISpfxReadWriteOperationsWebPartProps> {

  public render(): void {    
    const element: React.ReactElement<ISpfxReadWriteOperationsProps> = React.createElement(
      SpfxReadWriteOperations,
      {
        listID: this.properties.listID,
        context:this.context        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyFieldListPicker('listID', {
                  label: strings.ListNameFieldLabel,
                  selectedList: this.properties.listID,                  
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
