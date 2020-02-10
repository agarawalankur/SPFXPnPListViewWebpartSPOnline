import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
export interface IReactrenderListProps {
  description: string;  
  context: WebPartContext;
  siteURL: string;
  siteURLMultipe:string;
  listName: string;
  workflowName: string;
  fields: string;  
  filter: string;
  displayfields: string;
  title: string;
  groupByField: string;
  showFilter:boolean;
  compactView:boolean;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  
}
export interface IReactrenderListState {
  items: any[];
  ListDataArray:any[];
  Name:string;
  UserId: string;
  EmailId: string;
  Groupby:boolean;
  ViewFielsArray:IViewField[];
  GroupbyFieldArray:IGrouping[];

}
