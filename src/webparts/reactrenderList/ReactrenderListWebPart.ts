import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
} from '@microsoft/sp-webpart-base';
import * as strings from 'ReactrenderListWebPartStrings';
import ReactrenderList from './components/ReactrenderList';
import ControlsTest from './components/ReactrenderList';
import { IReactrenderListProps } from './components/IReactrenderListProps';
//import { PropertyPaneCheckbox } from '@microsoft/sp-property-pane';

export interface IReactrenderListWebPartProps {
  description: string;
  siteURL: string;  
  siteURLMultipe:string;
  workflowName:string;
  listName: string;
  fields:string;
  filter: string;  
  groupByField:string;
  displayfields: string;
  showFilter:boolean;
  compactView:boolean;
}

export default class ReactrenderListWebPart extends BaseClientSideWebPart<IReactrenderListWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactrenderListProps > = React.createElement(
      ReactrenderList,
      {
        description: this.properties.description,  
        context: this.context,
        siteURL: this.properties.siteURL,  
        siteURLMultipe:this.properties.siteURLMultipe,
        workflowName:this.properties.workflowName,        
        listName: this.properties.listName,
        fields: this.properties.fields,
        filter: this.properties.filter,
        title: this.properties.description,
        groupByField:this.properties.groupByField,
        showFilter:this.properties.showFilter,
        compactView:this.properties.compactView,
        displayfields:this.properties.displayfields,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.description = value;
        }    
      },
     
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
               /* PropertyPaneTextField('siteURL', {
                  label: strings.siteURLFieldLabel,
                  description:'To Get Current User Id',
                  disabled:true
                  
                }),*/
                PropertyPaneTextField('siteURLMultipe', {
                  label: strings.siteURLMultipeFieldLabel,                  
                  multiline:true                  
                }),                
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  description:'Same order as site URL'
                }),
                PropertyPaneTextField('workflowName', {
                  label: strings.workflowNameFieldLabel,
                  description:'Same order as List Name'
                }),
                PropertyPaneTextField('fields', {
                  label: strings.FieldsFieldLabel,
                  multiline: true,
                  description:'REST API Format with support of expand keyword Ex. ID,Title,LinkFilename,EncodedAbsUrl'
                }),
                PropertyPaneTextField('displayfields', {
                  label: strings.DisplayfieldsLabel,
                  multiline: true,
                  description:'Custom Format with Edit, View, File for item Links, Display as [User], [Date], [DateTime], [Lookup-Title] Ex. Edit,Title,File,Modified[DateTime]'
                }),               
                PropertyPaneTextField('groupByField', {
                  label: strings.GroupByfieldsLabel,
                  description:'[Single] - Ex WorkflowName'
                }),                
                PropertyPaneTextField('filter', {
                  label: strings.FilterFieldLabel,
                  multiline: true,
                  description:'REST API Format with [Today] & [Me] [Same site collection only], Use Id Suffix for User column like AssignedToId'
                }),
                PropertyPaneCheckbox('showFilter', {                  
                  text:'Show Filter'
                }),
                PropertyPaneCheckbox('compactView', {                  
                  text:'Compact View'
                }),
                
              ]
            }
          ]
        }
      ]
    };
  }
}
