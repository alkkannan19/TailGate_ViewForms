import * as React from 'react';
import styles from './TaskActivity.module.scss';
import { ITaskActivityProps } from './ITaskActivityProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { DetailsList, DetailsRow, IDetailsListProps, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { createListItems, IExampleItem } from '@uifabric/example-data';
import { fetch as fetchPolyfill } from "whatwg-fetch";
import pnp from "sp-pnp-js";
//import { CurrentUser } from "@pnp/sp/src/siteusers";//npm install @pnp/sp@1.3.8 @pnp/odata@1.3.8 @pnp/logging@1.3.8 @pnp/common@1.3.8 
//import {Web} from "sp-pnp-js";
import { sp } from "@pnp/sp";
//import {  ICamlQuery, Web } from "@pnp/sp/presets/all";
import { Web } from 'sp-pnp-js';
///pnpjs imports
export default class TaskActivity extends React.Component<ITaskActivityProps, {}> {
  public state;
  listName: any = "TailgateTasksActivity";
  private _items: IExampleItem[];
  private _columns: IColumn[];
  public curretSiteURL = new Web(this.props.context.pageContext.web.absoluteUrl);
  constructor(props) {
    super(props);
    this._columns = [
      {
        key: "column1",
        name: "Process Type",
        fieldName: "key",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      },
      {
        key: "column2",
        name: "Task Identifier",
        fieldName: "value",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      },
      {
        key: "column3",
        name: "Requester Name",
        fieldName: "name",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      },
      {
        key: "column4",
        name: "Request Date ",
        fieldName: "dateval",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      }
    ];
    this.state = {
      get_Network_Role: [],
      get_draftDetails: [],
      get_completeDetails: [],
      get_readonlyDetails: [],
      filterTaskDetails:"",
      filter_draftDetails:"",
      filter_completeDetails:"",
      filter_readonlyDetails:"",
    }
    this.getallDatas();
    this.getallDraftDetails();
    this.getallComleteDetails();
    this.getallReadOnlyDetails();
  }
  public componentDidMount() {
    console.log("componentDidMount");
  }
  public componentWillMount() {
    console.log("componentWillMount");
    sp.setup({
      spfxContext: this.props.context
    });
  }

  private getallDatas() {
    console.log("getallDatas");
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Submit'").get().then((Items: any) => {
          var _allItems: any[] = [];
          for (var i = 0; i < Items.length; i++) {
            var arritems = {
              key: Items[i]["ProcessType0"],
              value: Items[i]["TaskIdentifier"],
              name: Items[i]["Approvers"][0]["Title"],
              dateval: Items[i]["RequestDate"]
            };
            _allItems.push(arritems);
          }
          this.setState({ get_Network_Role: _allItems });
        });
    });
  }
  private getallDraftDetails() {
    console.log("getallDraftDetails");
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Draft'").get().then((Items: any) => {
          var _allItems: any[] = [];
          for (var i = 0; i < Items.length; i++) {
            var arritems = {
              key: Items[i]["ProcessType0"],
              value: Items[i]["TaskIdentifier"],
              name: Items[i]["Approvers"][0]["Title"],
              dateval: Items[i]["RequestDate"]
            };
            _allItems.push(arritems);
          }
          this.setState({ get_draftDetails: _allItems });
        });
    });
  }

  private getallComleteDetails() {
    console.log("getallDraftDetails");
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Completed' or RequesterNameId eq '" + UserId['Id'] + "' and Status eq 'Completed' or  SignoffsId eq '" + UserId['Id'] + "' and Status eq 'Completed'").get().then((Items: any) => {
          var _allItems: any[] = [];
          for (var i = 0; i < Items.length; i++) {
            var arritems = {
              key: Items[i]["ProcessType0"],
              value: Items[i]["TaskIdentifier"],
              name: Items[i]["Approvers"][0]["Title"],
              dateval: Items[i]["RequestDate"]
            };
            _allItems.push(arritems);
          }
          this.setState({ get_completeDetails: _allItems });
        });
    });
  }
  private getallReadOnlyDetails() {
    console.log("getallDraftDetails");
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Pending' or RequesterNameId eq '" + UserId['Id'] + "' and Status eq 'Submit' or SignoffsId eq '" + UserId['Id'] + "' and Status eq 'Pending'").get().then((Items: any) => {
          var _allItems: any[] = [];
          for (var i = 0; i < Items.length; i++) {
            var arritems = {
              key: Items[i]["ProcessType0"],
              value: Items[i]["TaskIdentifier"],
              name: Items[i]["Approvers"][0]["Title"],
              dateval: Items[i]["RequestDate"]
            };
            _allItems.push(arritems);
          }
          this.setState({ get_readonlyDetails: _allItems });
        });
    });
  }

  public render(): React.ReactElement<null> {
    return (
      <div>
        <div className={styles.container}>
          <h2>My Tasks</h2>
          <div className={styles.row}>
            <hr></hr>
            <Pivot>
              <PivotItem linkText="Active Tasks">
                <hr></hr>
                <DetailsList
                  items={this.state.get_Network_Role}
                  setKey="set"
                  columns={this._columns}
                  checkButtonAriaLabel="Row checkbox"
                />
              </PivotItem>
              <PivotItem linkText="Draft Requests">
                <hr></hr>
                <DetailsList
                  items={this.state.get_draftDetails}
                  setKey="set"
                  columns={this._columns}
                  checkButtonAriaLabel="Row checkbox"
                />
              </PivotItem>
              <PivotItem linkText="Completed Tasks">
                <hr></hr>
                <DetailsList
                  items={this.state.get_completeDetails}
                  setKey="set"
                  columns={this._columns}
                  checkButtonAriaLabel="Row checkbox"
                />
              </PivotItem>
              <PivotItem linkText="Read Only Tasks">
                <hr></hr>
                <DetailsList
                  items={this.state.get_readonlyDetails}
                  setKey="set"
                  columns={this._columns}
                  checkButtonAriaLabel="Row checkbox"
                />
              </PivotItem>
            </Pivot>
          </div></div>
      </div>
    );
  }
}
