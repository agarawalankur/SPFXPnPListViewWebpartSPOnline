import * as React from 'react';
import styles from './ReactrenderList.module.scss';
import { IReactrenderListProps, IReactrenderListState } from './IReactrenderListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from '@pnp/sp';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

import { SPHttpClient } from '@microsoft/sp-http';
import * as $ from 'jquery';

export default class ReactrenderList extends React.Component<IReactrenderListProps, IReactrenderListState> {
  constructor(props: IReactrenderListProps) {
    super(props);
    this.state = {
      items: [],
      Name: '',
      UserId: '',
      EmailId: '',
      Groupby: false,
      GroupbyFieldArray: [],
      ViewFielsArray: [],
      ListDataArray: [],
    };
  }
  public componentDidMount() {
    //update
    //sp.web.currentUser.get().then((user) => { 
    var checkifMultipleSitesArray = [];
    var ListNamesArray = [];
    var workflowNamesArray = [];
    const checkifMultipleSites = `${this.props.siteURLMultipe}`;
    if (checkifMultipleSites) {
      checkifMultipleSitesArray = checkifMultipleSites.split(',');
    }
    var ListNames = `${this.props.listName}`;
    ListNamesArray = ListNames.split(',');

    var workflowNames = `${this.props.workflowName}`;
    workflowNamesArray = workflowNames.split(',');

    checkifMultipleSitesArray.map((Urlvalue, indexVal) => {
      this.getCurrentUserWithAsyncAwait(Urlvalue)
        .then(result => {
          var CurrentUserId = result;
          var siteURL = Urlvalue;
          var restApi = siteURL + `/_api/web/lists/GetByTitle('` + ListNamesArray[indexVal] + `')/items?$top=5000`;
          var restApiListForm = siteURL + `/_api/web/lists/GetByTitle('` + ListNamesArray[indexVal] + `')/Forms?&select=ServerRelativeUrl,DecodedUrl&$filter=FormType eq 4`;

          if (`${this.props.fields}`) {
            var Fields = `${this.props.fields}`;
            restApi += "&$select=" + Fields;
          }
          else {
            restApi += "&$select=*";
          }
          // Check if filter provided 
          if (`${this.props.filter}`) {
            var filterval = `${this.props.filter}`;
            //------------------Date Filter  ----------------------------------------------------

            if (filterval.indexOf("[Today]") !== -1) {
              var todayDate = new Date();
              todayDate.setDate(todayDate.getDate() - 1);
              //setting zero hours
              var TodayISO = todayDate.toISOString();
              TodayISO = TodayISO.split("T")[0];
              TodayISO = "'" + TodayISO + "T20:00:00Z" + "'";
              filterval = this.replaceAll(filterval, "[Today]", TodayISO);
            }
            //------------------End Date Filter  -------------------------------------------------
            //--------------------------Current User Filter
            if (filterval.indexOf("[Me]") !== -1) {
              filterval = this.replaceAll(filterval, "[Me]", CurrentUserId);
            }
            //--------------------------End Current User Filter ----------------------------
            restApi += "&$filter=" + filterval;
          }
          if (`${this.props.displayfields}`) {
            const ViewFielsArray: IViewField[] = [];
            // dynamic view fields array
            var displayfields = `${this.props.displayfields}`;
            var DispFieldsArray = displayfields.split(',');
            DispFieldsArray.map((value) => {
              //DispFieldsArray.forEach(function (value) {
              //console.log(value);
              var data: IViewField = {
                name: "",
                displayName: "",
                sorting: true,
                maxWidth: 100,
                isResizable: true
              };
              if (value == "Edit") {
                data.name = 'ID',
                  data.displayName = 'Action',
                  data.sorting = true,
                  data.maxWidth = 60,
                  data.isResizable = false,
                  data.render = item => (
                    <a href={item.EditLink} target="_blank">Click here</a>
                  );
              }
              else if (value == "View") {
                data.name = 'ID',
                  data.displayName = 'View',
                  data.sorting = true,
                  data.maxWidth = 60,
                  data.isResizable = false,
                  data.render = item => (
                    <a href={item.ViewLink} target="_blank">Click here</a>
                  );
              }
              else if (value == "File") {
                data.name = 'LinkFilename',
                  data.displayName = 'File Name',
                  data.sorting = true,
                  data.maxWidth = 200,
                  data.isResizable = true,
                  data.render = item => (
                    <a href={item.EncodedAbsUrl} target="_blank">{item.LinkFilename}</a>
                  );
              }
              else if (value.indexOf("[Date]") !== -1) {
                var valueArrayDate = value.split("[");
                value = valueArrayDate[0];
                data.name = value,
                  data.displayName = value,
                  data.sorting = true,
                  data.maxWidth = 70,
                  data.isResizable = true,
                  data.render = (item) => {
                    const localizedEndDate = new Date(item[value]);
                    return (<span>{localizedEndDate.toLocaleString().split(',')[0]}</span>);
                  };
              }
              else if (value.indexOf("[DateTime]") !== -1) {
                var valueArrayDateTime = value.split("[");
                value = valueArrayDateTime[0];
                  data.name = value,
                  data.displayName = value,
                  data.sorting = true,
                  data.maxWidth = 100,
                  data.isResizable = true,
                  data.render = (item) => {
                    const localizedEndDate = new Date(item[value]);
                    return (<span>{localizedEndDate.toLocaleString()}</span>);
                  };
              }
              else if (value.indexOf("[User") !== -1) {
                  var valueArrayUser = value.split("[");
                  var Uservalue = valueArrayUser[0].split('-')[0];
                  var ColUservalue = valueArrayUser[1].split('-')[1].split(']')[0];                 
                  data.name = Uservalue + "." + ColUservalue;                 
                  data.displayName = Uservalue,
                  data.sorting = true,
                  data.maxWidth = 100,
                  data.isResizable = true,
                  data.render = (item) => {
                    var UserString = Uservalue + "." + ColUservalue; 
                    const UserName = item[UserString];
                    //console.log(item[value['Title']])
                    return (<span>{UserName}</span>);
                  };
              }
              else if (value.indexOf("[Lookup") !== -1) {
                var valueArrayLookup = value.split("[");
                var Lookupvalue = valueArrayLookup[0].split('-')[0];
                var Colvalue = valueArrayLookup[1].split('-')[1].split(']')[0];
                data.name = Lookupvalue + "." + Colvalue,
                  data.displayName = Lookupvalue,
                  data.sorting = true,
                  data.maxWidth = 100,
                  data.isResizable = true,
                  data.render = (item) => {
                    var UserString = Lookupvalue + "." + Colvalue;
                    const LookupItem = item[UserString];
                    //console.log(item[value['Title']])
                    return (<span>{LookupItem}</span>);
                  };
              }
              else {
                data.name = value,
                  data.displayName = value,
                  data.sorting = true,
                  data.maxWidth = 100,
                  data.isResizable = true;
              }
              ViewFielsArray.push(data);
            });
            this.setState({
              ViewFielsArray: ViewFielsArray
            });
          }
          else {
            const ViewFielsArrayDefault = [
              {
                name: 'ID',
                displayName: 'Action',
                sorting: true,
                maxWidth: 70,
                isResizable: false,
                render: item => (
                  <a href={item.EditLink} target="_blank">Click here</a>
                )
              },
              {
                name: 'Title',
                displayName: 'Title',
                sorting: true,
                maxWidth: 150,
                isResizable: true
              },
              {
                name: 'Modified',
                displayName: 'Modified',
                sorting: true,
                maxWidth: 100,
                isResizable: true
              }
            ];
            this.setState({
              ViewFielsArray: ViewFielsArrayDefault
            });

          }
          if (`${this.props.groupByField}`) {
            var GroupbyFieldValue = `${this.props.groupByField}`;

            const groupByFields: IGrouping[] = [
              {
                name: GroupbyFieldValue,
                order: GroupOrder.ascending                
              },
            ];
            this.setState({
              Groupby: true,
              GroupbyFieldArray: groupByFields
            });
          }
          console.log(restApi);
          var DisplayFormURl = "";
          var EditFormURl = "";
          this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
            .then(resp => { return resp.json(); })
            .then(items => {
              var resultsList = items.value;
              // Get Display form URL to create Link
              this.props.context.spHttpClient.get(restApiListForm, SPHttpClient.configurations.v1)
                .then(respForm => { return respForm.json(); })
                .then(Formitems => {
                  DisplayFormURl = Formitems.value[0].ResourcePath.DecodedUrl;
                  EditFormURl = DisplayFormURl.replace('DispForm', 'EditForm');

                  resultsList.map((obj, key) => {
                    obj.WorkflowName = workflowNamesArray[indexVal];
                    //obj.EditLink=Urlvalue + `/Lists/` + ListNamesArray[indexVal] + `/EditForm.aspx?ID=` + obj.Id;
                    obj.EditLink = Urlvalue + "/" + EditFormURl + `?ID=` + obj.Id;
                    obj.ViewLink = Urlvalue + "/" + DisplayFormURl + `?ID=` + obj.Id;
                  });
                  this.setState({
                    ListDataArray: this.state.ListDataArray.concat(resultsList)
                  });
                  this.setState({
                    items: this.state.ListDataArray
                  });
                });
              /*this.setState({
                items: ListDataArray ? ListDataArray : []
              }); */
            });
        });

    });
  }
  public render(): React.ReactElement<IReactrenderListProps> {
    const viewFields: IViewField[] = this.state.ViewFielsArray;
    if (this.state.Groupby) {
      return (
        <div className={styles.reactrenderList}>
          <div className={styles.container}>
            <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.description}
              updateProperty={this.props.updateProperty} />
            <ListView
              items={this.state.items}
              viewFields={viewFields}
              iconFieldName="EncodedAbsUrl"
              compact={this.props.compactView}
              showFilter={this.props.showFilter}
              selectionMode={SelectionMode.none}
              selection={this._getSelection}
              filterPlaceHolder="Search..."
              groupByFields={this.state.GroupbyFieldArray} />

          </div>
        </div>

      );
    }
    else {
      return (
        <div className={styles.reactrenderList}>
          <div className={styles.container}>
            <WebPartTitle displayMode={this.props.displayMode}
              title={this.props.description}
              updateProperty={this.props.updateProperty} />
            <ListView
              items={this.state.items}
              viewFields={viewFields}
              iconFieldName="EncodedAbsUrl"
              compact={this.props.compactView}
              showFilter={this.props.showFilter}
              selectionMode={SelectionMode.none}
              selection={this._getSelection}
              filterPlaceHolder="Search..."
            />
          </div>
        </div>

      );
    }
  }
  private _getSelection(items: any[]) {
    console.log('Selected items:', items);
  }
  private escapeRegExp(str) {
    return str.replace(/([.*+?^=!:${}()|\[\]\/\\])/g, "\\$1");
  }
  private replaceAll(str, find, replace) {
    return str.replace(new RegExp(this.escapeRegExp(find), 'g'), replace);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  public async getUserId(email: string): Promise<any> {
    await sp.site.rootWeb.ensureUser(email).then(result => {
      return result.data.Id;
    });

  }
  private async getCurrentUserWithAsyncAwait(siteurl) {
    const response = await this.props.context.spHttpClient.get(siteurl + '/_api/web/currentuser', SPHttpClient.configurations.v1);
    const user = await response.json();
    console.log(user.Id);
    return user.Id;
  }
  public GetUserDetails(Siteurl): Promise<any[]> {
    let url: string = Siteurl + `/_api/web/currentuser`;
    var items;
    $.ajax({
      url: url,
      headers: {
        Accept: "application/json;odata=verbose"
      },
      async: false,
    }).then((results): void => {
      items = results.d;

    });
    return items;
  }
  public formatDate(date) {
    var month_names = ["Jan", "Feb", "Mar",
      "Apr", "May", "Jun",
      "Jul", "Aug", "Sep",
      "Oct", "Nov", "Dec"];
    var dateval = new Date(date);
    var year = dateval.getFullYear();
    var month = dateval.getMonth();
    var dt = dateval.getDate();
    var dtString;
    var monthString;
    if (dt < 10) {
      dtString = '0' + dt;
    }
    if (month < 10) {
      monthString = '0' + month;
    }
    console.log(dtString + '-' + month_names[month] + '-' + year);

  }
}
