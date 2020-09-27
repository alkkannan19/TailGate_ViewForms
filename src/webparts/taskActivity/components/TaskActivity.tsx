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
import { IconButton, PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react';
//import { CurrentUser } from "@pnp/sp/src/siteusers";//npm install @pnp/sp@1.3.8 @pnp/odata@1.3.8 @pnp/logging@1.3.8 @pnp/common@1.3.8 
//import {Web} from "sp-pnp-js";
import { sp } from "@pnp/sp";
//import {  ICamlQuery, Web } from "@pnp/sp/presets/all";
import { Web } from 'sp-pnp-js';
///pnpjs imports
// import {
//    PeoplePicker,
//    PrincipalType
//  } from "@pnp/spfx-controls-react/lib/PeoplePicker";
export interface EditFormState {
  isTaskView: boolean,
  isEditView: boolean,
  isApproveView:boolean
}
export default class TaskActivity extends React.Component<ITaskActivityProps, {}> {
  public state;
  listName: any = "TailgateTasksActivity";
  private _items: IExampleItem[];
  private _columns: IColumn[];
  private draft_columns: IColumn[];
  options: IChoiceGroupOption[]
  public curretSiteURL = new Web(this.props.context.pageContext.web.absoluteUrl);
  constructor(props) {
    super(props);
    this.options = [
      { key: 'A', text: 'Approve' },
      { key: 'B', text: 'Return' }
    ];
    this.draft_columns = [];
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
        fieldName: "mode",
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
      },
    ];
    this.state = {
      get_activeDetails: [],
      get_draftDetails: [],
      get_completeDetails: [],
      get_readonlyDetails: [],
      filterTaskDetails: "",
      filter_draftDetails: "",
      filter_completeDetails: "",
      filter_readonlyDetails: "",
      isTaskView: true,
      isEditView: false,
      isApproveView:false,
      listItemId: "",
      topic: "",
      description: "",
      attachments: [],
      attachfileName: "",
      approveStatus: "Approve",
      comments: "",
      errorcomments: ""
    }
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._getSignOffPeoplePickerItems = this._getSignOffPeoplePickerItems.bind(this);
    this.getallActiveTasks();
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
  private _getPeoplePickerItems(items: any[]) {
    this.setState({
      allpeoplePicker_User: items,
      errorapproverUsers: false,
    });
  }
  private _getSignOffPeoplePickerItems(items: any[]) {
    this.setState({ allpeoplePicker2_User: items, errorSignoffUsers: false });
  }

  private getallActiveTasks() {

    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Submit'").get().then((ActiveItems: any) => {
          var _allItems: any[] = [];
          let modeObj: any;
          for (var i = 0; i < ActiveItems.length; i++) {
            let index = 0 + i;
            modeObj = <Link
              onClick={() =>
                this.getApproveForm(ActiveItems[index]["Id"])
              } href={""}>{ActiveItems[index]["TaskIdentifier"]} </Link>
            var arritems = {
              key: ActiveItems[i]["ProcessType0"],
              mode: modeObj,
              name: ActiveItems[i]["Approvers"][0]["Title"],
              dateval: ActiveItems[i]["RequestDate"]
            };
            _allItems.push(arritems);
          }
          this.setState({ get_activeDetails: _allItems });
        });
    });
  }
  private getallDraftDetails() {
    console.log("getallDraftDetails");
    this.draft_columns = [
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
        fieldName: "mode",
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
      },
      {
        key: "column5",
        name: "Edit",
        fieldName: "edit",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true
      }

    ];

    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Draft'").get().then((DraftItems: any) => {
          var _allItems: any[] = [];
          let editObj: any;
          for (var i = 0; i < DraftItems.length; i++) {
            let index = 0 + i;
            editObj = <IconButton
              style={{ color: "red" }}
              iconProps={{ iconName: "WindowEdit" }}//WindowEdit           
              onClick={() =>
                this.btnEditForm(DraftItems[index]["Id"])
              }
              title="Edit"
              ariaLabel="Add"
            />;
            var arritems = {
              key: DraftItems[i]["ProcessType0"],
              mode: DraftItems[i]["TaskIdentifier"],
              name: DraftItems[i]["Approvers"][0]["Title"],
              dateval: DraftItems[i]["RequestDate"],
              edit: editObj
            };
            _allItems.push(arritems);
          }
          this.setState({ get_draftDetails: _allItems });
        });
    });
  }
  private btnEditForm(id: any) {
    this.setState({ isTaskView: false })
    console.log(id);
    this.setState({ isTaskView: false });
    sp.web.lists.getByTitle(this.listName).items.getById(id).get().then((singleItem: any) => {
      console.log(singleItem);
    });
  }
  private getApproveForm(id: any) {
    this.setState({ listItemId: id })
    var url;
    this.setState({ isTaskView: false });
    sp.web.lists.getByTitle(this.listName).items.getById(id).get().then((singleItem: any) => {
      console.log(singleItem);
      var _allItems: any[] = [];
      let item = sp.web.lists.getByTitle(this.listName).items.getById(id).attachmentFiles.get()
        .then(v => {
          for (var i = 0; i < v.length; i++) {
            url = <Link href={this.props.context.pageContext.web.absoluteUrl +
              `/Lists/DiscussionList/Attachments/` +
              id +
              `/` +
              v[i].FileName}>{v[i].FileName}</Link>
            _allItems.push(url);
          }
          console.log(url);
          this.setState({
            attachments: _allItems
          });

        });

      this.setState({
        topic: singleItem.Topic,
        description: singleItem.Description,

      })
    });
  }

  private getallComleteDetails() {
    console.log("getallComleteDetails");
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Completed' or RequesterNameId eq '" + UserId['Id'] + "' and Status eq 'Completed' or  SignoffsId eq '" + UserId['Id'] + "' and Status eq 'Completed'").get().then((Items: any) => {
          var _allItems: any[] = [];
          for (var i = 0; i < Items.length; i++) {
            var arritems = {
              key: Items[i]["ProcessType0"],
              mode: Items[i]["TaskIdentifier"],
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
    console.log("getallReadOnlyDetails");
    this.curretSiteURL.currentUser.get().then((UserId) => {
      console.log("Current User Id " + UserId['Id'] + " Current User Name " + UserId['Title']);
      sp.web.lists.getByTitle(this.listName).items.select("*,Approvers/Name,Approvers/Title").expand("Approvers")
        .filter("ApproversId eq '" + UserId['Id'] + "' and Status eq 'Pending' or RequesterNameId eq '" + UserId['Id'] + "' and Status eq 'Submit' or SignoffsId eq '" + UserId['Id'] + "' and Status eq 'Pending'").get().then((AllItems: any[]) => {
          var _allItems: any[] = [];

          for (var i = 0; i < AllItems.length; i++) {

            var arritems = {
              key: AllItems[i]["ProcessType0"],
              mode: AllItems[i]["TaskIdentifier"],
              name: AllItems[i]["Approvers"][0]["Title"],
              dateval: AllItems[i]["RequestDate"],

            };
            _allItems.push(arritems);
          }
          this.setState({ get_readonlyDetails: _allItems });
        });
    });
  }
  public _alertClicked = (): void => {
    this.setState({
      isTaskView: true
    });
  }
  fileUploadCallback = event => {
    const file = event.target.files[0];
    this.setState({
      fileAttach: file,
      errorfileAttach: ""
    });
  };
  public handleLinkClick = (item: PivotItem) => {
    console.log(item.props.itemKey);
    if (item.props.itemKey == "2") {
      this.getallDraftDetails();
    }
  };
  private cancel = (): void => {
    alert("Cancel");
  }

  private submitForm = (): void => {
    this.state.topicValue.trim().length > 0 ? "" : this.setState({ errortopicValue: "Topic is required" });
    this.state.descriptionValue.trim().length > 0 ? "" : this.setState({ errordescriptionValue: "Description is required" });
    this.state.fileAttach.length == 0 ? this.setState({ errorfileAttach: "Approvers is required" }) : "";
    this.state.allpeoplePicker_User.length > 0 ? "" : this.setState({ errorapproverUsers: true });
    this.state.allpeoplePicker2_User.length > 0 ? "" : this.setState({ errorSignoffUsers: true });

    if (this.state.topicValue.trim().length > 0 && this.state.descriptionValue.trim().length > 0 && this.state.fileAttach && this.state.allpeoplePicker2_User.length > 0) {
      var today=today.getDay()+"/"+(today.getMonth()+1)+"/"+today.getFullYear();
     // let today = new Date().toISOString().slice(0, 10);
      sp.web.lists.getByTitle("TailgateTasksActivity").items.add({
        Topic: this.state.topicValue,
        Description: this.state.descriptionValue,
        Status: "Submit",
        ProcessType0: "Tailgate",
        TaskIdentifier: "Tailgate Topic",
        RequestDate: today,
        RequesterNameId: {
          results: [this.state.currentUserId]// User/Groups ids as an array of numbers
        },
        ApproversId: {
          results: [this.state.allpeoplePicker_User.length > 0 ? this.state.allpeoplePicker_User[0]["id"] : []]  // User/Groups ids as an array of numbers
        },

        SignoffsId: {
          results: [this.state.allpeoplePicker2_User[0]["id"]]  // User/Groups ids as an array of numbers
        },
      })
        .then((disID: any) => {
          //   console.log("Add Items to List SuccessFully");     
          //   console.log(disID.data.Id);
          //   let item = sp.web.lists
          //   .getByTitle("TailgateTasksActivity")
          //   .items.getById(disID.data.Id)
          //   item.attachmentFiles.add("Test", this.state.fileAttach).then(result => {   
          //   console.log("File uploaded successfully...")   
          // }); 
          let item = sp.web.lists.getByTitle("TailgateTasksActivity").items.getById(disID.data.Id);
          item.attachmentFiles.add(this.state.fileAttach.fileName, this.state.fileAttach).then(v => {
            console.log("File upload successfully...!");
            alert("Submitted Successfully..!");
            this.setState({
              topicValue: "",
              descriptionValue: "",
              fileAttach: [""],
              allpeoplePicker_User: [""],
              allpeoplePicker2_User: [""]
            })
          });

        });
    }

  }
  public _alertClicked1 = (): void => {
    if (this.state.comments.trim().length > 0) {
      sp.web.lists.getByTitle(this.listName).items.getById(this.state.listItemId).update({
        Status: this.state.approveStatus,
        Comments: this.state.comments
      }).then(s => {
        console.log("Items updated Successfully");
        alert("Task updated successfully...!!!");
        this.setState({ isTaskView: true });
      });
    } else {
      this.setState({ errorcomments: "Comments is Required" });
    }

  }
  public _onRenderListView(
  ): JSX.Element {
    return (<div><div>
      <div className={styles.container}>
        <h2>My Tasks</h2>
        <div className={styles.row}>
          <hr></hr>
          <Pivot onLinkClick={this.handleLinkClick}>
            <PivotItem linkText="Active Tasks" itemKey="1">
              <hr></hr>
              <DetailsList
                items={this.state.get_activeDetails}
                setKey="set"
                columns={this._columns}
                checkButtonAriaLabel="Row checkbox"
              />
            </PivotItem>
            <PivotItem linkText="Draft Requests" itemKey="2">
              <hr></hr>
              <DetailsList
                items={this.state.get_draftDetails}
                setKey="set"
                columns={this.draft_columns}
                checkButtonAriaLabel="Row checkbox"
              />
            </PivotItem>
            <PivotItem linkText="Completed Tasks" itemKey="3">
              <hr></hr>
              <DetailsList
                items={this.state.get_completeDetails}
                setKey="set"
                columns={this._columns}
                checkButtonAriaLabel="Row checkbox"
              />
            </PivotItem>
            <PivotItem linkText="Read Only Tasks" itemKey="4">
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
    </div> </div>)
  }
  public _onRenderApproveView(
  ): JSX.Element {
    return (<div className={styles.taskActivity}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <label className={styles.divalign}>Topic : </label>
          </div>
          <div className={styles.col_6}>
            <label>{this.state.topic}</label>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <label className={styles.divalign}>Description : </label>
          </div>
          <div className={styles.col_6}>
            <label>{this.state.description}</label>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <label className={styles.divalign}>Attachments : </label>
          </div>
          <div className={styles.col_6}>
            <label>{this.state.attachments}</label>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <label className={styles.divalign}>Action History</label>
            </div>
            <div className={styles.col_6}>
            <label>Action History</label>
          </div>
      </div>

        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <ChoiceGroup defaultSelectedKey="A" label="Action" options={this.options} onChange={this._onChange} required={true} />
          </div>

        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <TextField label="Comments" required value={this.state.comments} onChanged={newVal => {
              newVal && newVal.length > 0
                ? this.setState({
                  comments: newVal,
                  errorcomments: ""
                })
                : this.setState({
                  comments: newVal,
                });
            }} errorMessage={this.state.errorcomments} />
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_3}>
            <DefaultButton text="Cancel" onClick={this._alertClicked} />
          </div>
          <div className={styles.col_3}>
            <PrimaryButton text="Submit" onClick={this._alertClicked1} /></div>
        </div>
      </div>
  )
  }

  public _onRenderEditView(
  ): JSX.Element {
    return (
      <div className={styles.container}>
        <hr></hr>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <TextField label="Topic" required
              value={this.state.topicValue}
              onChanged={newVal => {
                newVal && newVal.length > 0
                  ? this.setState({
                    topicValue: newVal,
                    errortopicValue: ""
                  })
                  : this.setState({
                    topicValue: newVal,
                    errortopicValue:
                      "Topic is required"
                  });
              }}
              errorMessage={this.state.errortopicValue}
            />
          </div>
          <div className={styles.col_6}>
            <Label required>Approvers</Label>
             {/* <PeoplePicker
              context={this.props.context}
              titleText=""
              personSelectionLimit={1}
              groupName={""}
              showtooltip={false}
              // isRequired={true}
              disabled={false}
              ensureUser={true}
              selectedItems={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}

            />  */}
            {/* {this.state.errorapproverUsers ? <Label className={styles.pickerlabelErrormsg}>Approvers is required</Label> : ""} */}
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <TextField label="Description" required
              value={this.state.descriptionValue}
              onChanged={newDesVal => {
                newDesVal && newDesVal.length > 0
                  ? this.setState({
                    descriptionValue: newDesVal,
                    errordescriptionValue: ""
                  })
                  : this.setState({
                    descriptionValue: newDesVal,
                    errordescriptionValue:
                      "Description is required"
                  });
              }}
              multiline rows={3} errorMessage={this.state.errordescriptionValue} />
          </div>
          <div className={styles.col_6}>
            <Label required>Sign offs</Label>
             {/* <PeoplePicker
              //  peoplePickerCntrlclassName={styles.pickerErrormsg}
              context={this.props.context}
              titleText=""
              personSelectionLimit={1}
              groupName={""}
              showtooltip={false}
              //  isRequired={true}
              disabled={false}
              ensureUser={true}
              selectedItems={this._getSignOffPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            //errorMessage={this.state.SignoffUsers}
            />  */}
            {this.state.errorSignoffUsers ? <Label className={styles.pickerlabelErrormsg}>Sign Offs is required</Label> : ""}

          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_6}>
            <div>
              <Label required>Attachment</Label>
              <input type="file" multiple accept=".xlsx,.xls,.doc, image/*, .docx,.ppt, .pptx,.txt,.pdf" onChange={this.fileUploadCallback}
              />
              {this.state.errorfileAttach ? <Label className={styles.pickerlabelErrormsg}>Attachment is required</Label> : ""}
            </div>

          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.col_3}>
            <PrimaryButton className={styles.btnDraft} text="Save as Draft" onClick={this.cancel} />
          </div>
          <div className={styles.col_3}>
            <PrimaryButton className={styles.btnSubmit} text="Submit" onClick={this.submitForm} />
          </div>
        </div>
      </div>
    )

  }
  public _onChange(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
    console.dir(option);
    this.setState({ approveStatus: option.text })
  }

  public render(): React.ReactElement<null> {
    return (
      <div>
        {this.state.isTaskView ? this._onRenderListView()
          : this._onRenderApproveView()}
      </div>
    );
  }
}
