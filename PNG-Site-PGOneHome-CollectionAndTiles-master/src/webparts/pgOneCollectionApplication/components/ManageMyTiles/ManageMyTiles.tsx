import * as React from 'react';
import { Text, Dropdown, PrimaryButton, IDropdownStyles, IDropdownOption, DetailsList, IColumn, Link, Icon, DetailsListLayoutMode, SelectionMode, SearchBox, Panel, PanelType, TextField, Label, MessageBar, DefaultButton, MessageBarType, Dialog, DialogFooter, DialogType, Shimmer, ITextField, Toggle } from 'office-ui-fabric-react';
import { PnPHelper } from '../PnPHelper/PnPHelper';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from './ManageTile.module.scss';
import { SPPermission } from '@microsoft/sp-page-context';
import Pagination from 'office-ui-fabric-react-pagination';
import { sp } from '@pnp/sp';
import { stringIsNullOrEmpty } from '@pnp/common';
interface IManageMyTilesList {
  Title: string;
  ID?: string;
}

export interface IColorMasterList {
  Id?: number;
  Title?: string;
  BgColor?: string;
  ForeColor?: string;
}

interface IManageMyTilesState {
  status: string;
  resourceListItems: any[];
  myTilesItems: IManageMyTilesList[];
  currentItem: IManageMyTilesList[];
  columns: IColumn[];
  showViewPanel: boolean;
  showEditPanel: boolean;
  hideDeleteDialog: boolean;
  dissableRemoveButton: boolean;
  formDataIsChanged: boolean;
  isRecordIsAlreadyInApproval: boolean;
  valueRequiredErrorMessage: string;
  valueValidUrlLinkErrorMessage: string;
  valueValidEamilErrorMessage: string;
  valueMaxCharacterErrorMessage: string;
  TileId: number;
  TileTitle?: string;
  TileDescription: string;
  TileKewords: string;
  TileURLLink: string;
  TileOwner: string;
  TileApprovedDate?: Date;
  RequestedAction: string;
  TileColorCodeBgColor: string;
  TileColorCodeForeColor: string;
  TileColorCodeId: number;
  TileColorCategory: string;
  TileAvailableExternal: number;
  isDataLoaded: boolean;

  currentPage?: number;
  totalPage?: number;
  itemsPerPage?: number;

  tileRequestColorCode: number;
  lstColorMaster?: IColorMasterList[];

}

interface IManageMyTilesProps {
  wpContext: WebPartContext;
  resourceListItems: any[];
  callBackForRequestSection: any;
}

const ddNoOFItems = [
  { key: 10, text: '10' },
  { key: 20, text: '20' },
  { key: 30, text: '30' },
  { key: 50, text: '50' },
  { key: 100, text: '100' },
];
const ddOptionsStyles: Partial<IDropdownStyles> = { dropdown: { width: 100, height: 28, marginLeft: 6, marginRight: 6, marginTop: -7, marginBottom: -7 } };

export class ManageMyTiles extends React.Component<IManageMyTilesProps, IManageMyTilesState> {

  private pnpHelper: PnPHelper;
  private _allItems: IManageMyTilesList[];
  private allApplicationTiles: any;
  private allApplicationRequest: any;
  private lstTileRequest = "ApplicationRequests";
  private ApplicationMasterListName: string = "ApplicationMaster";
  private itemsForPagination: IManageMyTilesList[];
  private selectedColorId: number = 1;
  private colorOptions: IDropdownOption[] = [{ key: 0, text: 'Select a category' }];
  private showError: boolean = false;
  private lstColorMaster = "ColorMaster";
  private ctrTitle: React.RefObject<ITextField>;
  private ctrDescription: React.RefObject<ITextField>;
  private ctrKeywords: React.RefObject<ITextField>;
  private ctrURL: React.RefObject<ITextField>;
  private ctrOwnerEmail: React.RefObject<ITextField>;

  private urlPattern = /(http(s)?:\/\/.)(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,63984}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&\/=]*)/;
  private emailPattern = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@(pg.com)$/;
  private isOwnerMember = this.props.wpContext.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);

  private errTitle: string = "PNG-Site-PGOneHome-CollectionAndTiles";
  private errModule: string = "ManageMyTiles.tsx";

  constructor(props: IManageMyTilesProps, state: IManageMyTilesState) {
    super(props);

    const columns: IColumn[] = [
      {
        key: 'ID', name: 'Sr. No', fieldName: 'ID',
        minWidth: 20, maxWidth: 60, isResizable: true, isCollapsible: true,
        //onColumnClick: this._onColumnClickSort.bind(this),
        onRender: (item, index) => (
          <div>
            {(this.state.itemsPerPage * this.state.currentPage + index + 1) - (this.state.itemsPerPage)}
          </div>
        ),
      },
      {
        key: 'Title', name: 'Title',
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClickSort.bind(this),
        isRowHeader: true,
        data: 'string',
        isPadded: true,
        fieldName: 'Title', minWidth: 100, isResizable: true, isCollapsible: false,

      },

      {
        key: 'IsActive', name: 'Status', fieldName: 'IsActive', minWidth: 80, maxWidth: 100, isResizable: true, isCollapsible: true,
        onRender: (item) => (
          <div style={{ color: "#323130", fontSize: 14 }}>
            {item.IsActive == true ? "Active" : "Inactive"}
          </div>
        ),
        onColumnClick: this._onColumnClickSort.bind(this),
        headerClassName: this.isOwnerMember ? "" : "optionalColumn",
        className: this.isOwnerMember ? "" : "optionalColumn"
      },
      {
        key: 'View', name: 'Action', fieldName: '', minWidth: 100, maxWidth: 150, isResizable: true, isCollapsible: false,
        onRender: (item) => (
          <div>
            <Link aria-label="View" title="View" onClick={() => this.showViewPanel(item)}><Icon className="icons" iconName="View" /></Link>
            {/* <Link aria-label="Edit" title="Edit" onClick={() => { this.showEditPanel(item); }} ><Icon iconName="Edit" /></Link>
            <Link aria-label="Delete" title="Delete" onClick={() => { this.deleteTile(item); }}><Icon iconName="Delete" /></Link> */}
          </div>
        ),
      }
    ];
    const columnsForAdmin: IColumn[] = [
      {
        key: 'ID', name: 'Sr. No', fieldName: 'ID',
        minWidth: 20, maxWidth: 60, isResizable: true, isCollapsible: true,
        //onColumnClick: this._onColumnClickSort.bind(this),
        onRender: (item, index) => (
          <div>
            {(this.state.itemsPerPage * this.state.currentPage + index + 1) - (this.state.itemsPerPage)}
          </div>
        ),
      },
      {
        key: 'Title', name: 'Title',
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClickSort.bind(this),
        isRowHeader: true,
        data: 'string',
        isPadded: true,
        fieldName: 'Title', minWidth: 100, isResizable: true, isCollapsible: false,

      },

      {
        key: 'IsActive', name: 'Status', fieldName: 'IsActive', minWidth: 80, maxWidth: 100, isResizable: true, isCollapsible: true,
        onRender: (item) => (
          <div style={{ color: "#323130", fontSize: 14 }}>
            {item.IsActive == true ? "Active" : "Inactive"}
          </div>
        ),
        onColumnClick: this._onColumnClickSort.bind(this),
        // headerClassName: this.isOwnerMember ? "" : "optionalColumn",
        className: this.isOwnerMember ? "" : "optionalColumn"
      },
      {
        key: 'View', name: 'Action', fieldName: '', minWidth: 100, maxWidth: 150, isResizable: true, isCollapsible: false,
        onRender: (item) => (
          <div>
            <Link aria-label="View" title="View" onClick={() => this.showViewPanel(item)}><Icon className="icons" iconName="View" /></Link>
            {/* <Link aria-label="Edit" title="Edit" onClick={() => { this.showEditPanel(item); }} ><Icon iconName="Edit" /></Link>
            <Link aria-label="Delete" title="Delete" onClick={() => { this.deleteTile(item); }}><Icon iconName="Delete" /></Link> */}
          </div>
        ),
      }
    ];
    const columnsForEndUser: IColumn[] = [
      {
        key: 'ID', name: 'Sr. No', fieldName: 'ID',
        minWidth: 20, maxWidth: 60, isResizable: true, isCollapsible: true,
        //onColumnClick: this._onColumnClickSort.bind(this),
        onRender: (item, index) => (
          <div>
            {(this.state.itemsPerPage * this.state.currentPage + index + 1) - (this.state.itemsPerPage)}
          </div>
        ),
      },
      {
        key: 'Title', name: 'Title',
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClickSort.bind(this),
        isRowHeader: true,
        data: 'string',
        isPadded: true,
        fieldName: 'Title', minWidth: 100, isResizable: true, isCollapsible: false,

      },
      {
        key: 'View', name: 'Action', fieldName: '', minWidth: 100, maxWidth: 150, isResizable: true, isCollapsible: false,
        onRender: (item) => (
          <div>
            <Link aria-label="View" title="View" onClick={() => this.showViewPanel(item)}><Icon className="icons" iconName="View" /></Link>
            {/* <Link aria-label="Edit" title="Edit" onClick={() => { this.showEditPanel(item); }} ><Icon iconName="Edit" /></Link>
            <Link aria-label="Delete" title="Delete" onClick={() => { this.deleteTile(item); }}><Icon iconName="Delete" /></Link> */}
          </div>
        ),
      }
    ];
    this.state = {
      status: "Ready",
      resourceListItems: this.props.resourceListItems,

      myTilesItems: [],
      columns: this.isOwnerMember ? columnsForAdmin : columnsForEndUser,
      currentItem: [],

      showViewPanel: false,
      showEditPanel: false,
      hideDeleteDialog: true,
      dissableRemoveButton: false,
      formDataIsChanged: false,
      isRecordIsAlreadyInApproval: true,

      valueRequiredErrorMessage: this.props.resourceListItems["required_field_validation_message"],
      valueValidUrlLinkErrorMessage: "",
      valueValidEamilErrorMessage: "",
      valueMaxCharacterErrorMessage: this.props.resourceListItems["validation_maximum_characters_text"],


      TileId: null,
      TileTitle: "",
      TileDescription: "",
      TileKewords: "",
      TileURLLink: "",
      TileOwner: "",
      TileColorCodeBgColor: "",
      TileColorCodeForeColor: "",
      TileColorCategory: "",
      TileColorCodeId: null,
      TileAvailableExternal: 1,
      RequestedAction: "",


      isDataLoaded: false,


      currentPage: 1,
      totalPage: 1,
      itemsPerPage: 10,

      //TileApprovedDate:Date,
      tileRequestColorCode: 1,
      lstColorMaster: []

    };
    this.pnpHelper = new PnPHelper(this.props.wpContext);

    this.ctrTitle = React.createRef();
    this.ctrDescription = React.createRef();
    this.ctrKeywords = React.createRef();
    this.ctrURL = React.createRef();
    this.ctrOwnerEmail = React.createRef();

  }



  public render(): React.ReactElement<IManageMyTilesProps> {
    return (
      <React.Fragment>
        <div>
          <div className={styles.manageMyTiles}>
            <div className={`ms-Grid`}>
              <div className="ms-Grid-row ">
                <div className="ms-Grid-col ms-sm10 ms-md12 ms-lg12 ms-xl12">
                  <h2 className="ms-font-m ms-fontWeight-semibold">{this.state.resourceListItems['manage_my_tiles_label']}</h2>
                </div>
              </div>
              <hr role="presentation"></hr>
              <div className={`ms-Grid-row ${styles.searchRow}`}>
                <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg3 ms-xl3 detailsListContoll">
                  <div className="paginationDropdown">
                    
                    <label id="lbl1">Show</label>
                    <Dropdown
                      // placeholder="Select an option"
                      label=""
                      options={ddNoOFItems}
                      styles={ddOptionsStyles}
                      selectedKey={this.state.itemsPerPage}
                      data-is-focusable={true}
                      onChange={this.ddItemsPerPageChanges.bind(this)}
                      aria-labelledby="lbl1 lbl2"
                     
                    />
                    <label id="lbl2">entries</label>
                    
                  </div>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg3 ms-xl3 searchBoxPG detailsListContoll">
                  <SearchBox placeholder={this.state.resourceListItems["review_search_placeholder_text"]} onSearch={this._onFilter.bind(this)} onChange={this._onFilter.bind(this)} />
                </div>
              </div>
              <div className={`ms-Grid-row ${styles.gridRowDetailsList}`}>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12" >
                  <DetailsList items={this.state.myTilesItems}
                    setKey="none"
                    ariaLabelForListHeader="Column headers. Click to sort."
                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                    ariaLabelForSelectionColumn="Toggle selection"
                    checkButtonAriaLabel="Row checkbox"
                    columns={this.state.columns}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                    selectionMode={SelectionMode.none}
                    className={styles.myTilesItems}
                  />
                  <Shimmer className={styles.shimmerLoader} isDataLoaded={this.state.isDataLoaded} />
                  <Shimmer className={styles.shimmerLoader} width="75%" isDataLoaded={this.state.isDataLoaded} />
                  <Shimmer className={styles.shimmerLoader} width="50%" isDataLoaded={this.state.isDataLoaded} />
                  {this.state.myTilesItems.length == 0 && this.state.isDataLoaded && (
                    <div> No records found.</div>
                  )
                  }
                  {this.state.myTilesItems.length > 0 && this.state.isDataLoaded && (
                    <div className="paginationList">
                      <Pagination currentPage={this.state.currentPage} totalPages={this.state.totalPage} onChange={(page: any) => { this.paginationChanged(page); }} />
                    </div>
                  )
                  }
                </div>
              </div>
            </div>
          </div>
          {/* View Tiles Panel */}
          <Panel
            isLightDismiss
            className={styles.ManageMyTilesPanel}
            headerText={this.state.resourceListItems["add_tile_manage_my_tiles_label"]}
            isOpen={this.state.showViewPanel}
            onDismiss={this.hideAllPanel}
            closeButtonAriaLabel="Close"
            type={PanelType.smallFixedFar}
            style={{ fontWeight: "bold" }}
            // onRenderFooterContent={this._onRenderTileEditFooterContent}
            isFooterAtBottom={true}
            headerClassName={styles.headerText}
          >
            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_title_label"]} </Label>
            <Label className={styles["ms-font-m"]}>{this.state.currentItem['Title'] == undefined ? "-" : this.state.currentItem['Title']}</Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_description_label"]}</Label>
            <Label className={styles["ms-font-m"]}>{this.state.currentItem['Description'] == undefined ? "-" : this.state.currentItem['Description']}</Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_keywords_label"]}</Label>
            <Label className={styles["ms-font-m"]}>{this.state.currentItem['SearchKeywords'] == undefined ? "-" : this.state.currentItem['SearchKeywords']}</Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_url_link_label"]}</Label>
            <Label className={styles["ms-font-m"]} >
              {this.state.currentItem['Link']}<br></br>
              <Link data-interception="off" href={this.state.currentItem['Link'] == undefined ? "#" : this.state.currentItem['Link']} target="_blank">{this.state.resourceListItems['add_tile_try_link_label']}</Link>
            </Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_owneremail_label"]}</Label>
            <Label className={styles["ms-font-m"]}>{this.state.currentItem['OwnerEmail'] == undefined ? "-" : this.state.currentItem['OwnerEmail']}</Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_category_label"]}</Label>
            <Label className={styles["ms-font-m"]}>{this.state.TileColorCategory == undefined ? "-" : this.state.TileColorCategory}</Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_available_external_label"]}</Label>
            <Label className={styles["ms-font-m"]}>
              {this.state.currentItem['AvailableExternal'] == null ? "-" : this.state.currentItem['AvailableExternal'] == 0 ? "Yes" : "No"}
            </Label>

            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_approved_date_label"]}</Label>
            {/* new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "2-digit", hour: 'numeric', minute: 'numeric', }).format(new Date(this.state.currentItem['ApprovedDate']))} */}
            <Label className={styles["ms-font-m"]}>
              {this.state.currentItem['ApprovedDate'] == null ? "-" :
                new Date(this.state.currentItem['ApprovedDate'].toString()).toLocaleString()
              }
            </Label>


          </Panel>


          {/* Edit Tiles Panel */}
          <Panel
            isLightDismiss
            className={styles.ManageMyTilesEditPanel}
            headerText={this.state.resourceListItems["add_tile_manage_my_tiles_label"]}
            isOpen={this.state.showEditPanel}
            onDismiss={this.hideAllPanel}
            closeButtonAriaLabel="Close"
            type={PanelType.smallFixedFar}
            style={{ fontWeight: "bold" }}
            onRenderFooterContent={this._onRenderTileEditFooterContent}
            isFooterAtBottom={true}
            headerClassName={styles.headerText}

          >
            <TextField
              label={this.state.resourceListItems["add_tile_title_label"]}
              placeholder="Tile name"
              onChange={this._onTileTitleChange.bind(this)}
              required={true}
              validateOnLoad={false}

              value={this.state.TileTitle == undefined ? "-" : this.state.TileTitle}
              //errorMessage={this.state.TileTitle.trim() === "" ? this.state.valueRequiredErrorMessage : ""}
              onGetErrorMessage={this.getTileRequestTitleErrorMessage.bind(this)}
              //maxLength={255}

              validateOnFocusIn={true}
              validateOnFocusOut={true}
              componentRef={this.ctrTitle}

            />

            <TextField
              label={this.state.resourceListItems["add_tile_description_label"]}
              placeholder="Tile description"

              onChange={this._onTileDescriptionChange.bind(this)}
              required={true}
              validateOnLoad={false}

              value={this.state.TileDescription == undefined ? "-" : this.state.TileDescription}
              multiline
              rows={3}
              //errorMessage={this.state.TileDescription.trim() === "" ? this.state.valueRequiredErrorMessage : ""}
              onGetErrorMessage={this.getTileRequestDescriptionErrorMessage.bind(this)}
              validateOnFocusIn={true}
              validateOnFocusOut={true}
              componentRef={this.ctrDescription}
            />
            <TextField
              label={this.state.resourceListItems["add_tile_keywords_label"]}
              placeholder="Comma separated keywords. e.g. Home, Document"
              multiline
              rows={3}
              onChange={this._onTileKewordChange.bind(this)}
              required={true}

              validateOnFocusOut={true}
              value={this.state.TileKewords == undefined ? "-" : this.state.TileKewords}
              //errorMessage={this.state.TileKewords.trim() === "" ? this.state.valueRequiredErrorMessage : ""}
              onGetErrorMessage={this.getTileRequestKeywordsErrorMessage.bind(this)}
              validateOnLoad={false}
              validateOnFocusIn={true}
              componentRef={this.ctrKeywords}

            />

            <TextField
              label={this.state.resourceListItems["add_tile_url_link_label"]}
              placeholder="e.g. https://pgone.pg.com"

              onChange={this._onTileLinkURLChange.bind(this)}
              required={true}
              validateOnLoad={false}

              value={this.state.TileURLLink == undefined ? "#" : this.state.TileURLLink}
              //errorMessage={this.state.valueValidUrlLinkErrorMessage}
              rows={3}
              onGetErrorMessage={this.getTileRequestURLLinkErrorMessage.bind(this)}
              validateOnFocusIn={true}
              validateOnFocusOut={true}
              componentRef={this.ctrURL}
            />
            <Link
              className={styles["ms-font-m"]}
              href={this.state.TileURLLink == undefined ? "#" : this.state.TileURLLink}
              target="_blank"
              disabled={!this.urlPattern.test(this.state.TileURLLink)}
            >
              {this.state.resourceListItems["add_tile_try_link_label"]}
            </Link>

            {this.isOwnerMember
              ?
              <TextField
                label={this.state.resourceListItems["add_tile_owneremail_label"]}
                placeholder="e.g. email@pg.com"

                onChange={this._onTileEmailChange.bind(this)}
                required={true}
                validateOnLoad={false}

                value={this.state.TileOwner == undefined ? "-" : this.state.TileOwner}
                readOnly={false}
                //errorMessage={this.state.valueValidEamilErrorMessage}
                onGetErrorMessage={this.getTileRequestOwnerEmailErrorMessage.bind(this)}
                validateOnFocusIn={true}
                validateOnFocusOut={true}
                componentRef={this.ctrOwnerEmail}


              />
              : <div>
                <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_owneremail_label"]}</Label>
                <Label className={styles["ms-font-m"]}>{this.state.TileOwner == undefined ? "-" : this.state.TileOwner}</Label>
              </div>
            }

            <Dropdown
              options={this.colorOptions}
              selectedKey={this.state.TileColorCodeId}
              label={this.state.resourceListItems["add_tile_category_label"]}
              ariaLabel={this.state.resourceListItems["add_tile_category_label"]}
              required={true}
              onChange={this._onColorCodeChange.bind(this)}
              errorMessage={this.selectedColorId === 0 && this.showError && this.state.resourceListItems['required_field_validation_message']}
            ></Dropdown>


            <Toggle
              label={this.state.resourceListItems['add_tile_available_external_label']}
              onText="Yes"
              offText="No"
              checked={this.state.TileAvailableExternal == 1 ? false : true}
              onChange={this._onColRequestUndeletableChange.bind(this)}
            />


            <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_approved_date_label"]}</Label>
            <Label className={styles["ms-font-m"]}>
              {this.state.currentItem['ApprovedDate'] == null ? "-" :
                new Date(this.state.currentItem['ApprovedDate'].toString()).toLocaleString()
              }
            </Label>
          </Panel>

          <Dialog
            hidden={this.state.hideDeleteDialog}
            onDismiss={() => this.setState({ hideDeleteDialog: true })}
            dialogContentProps={{
              type: DialogType.normal,
              showCloseButton: false
            }}
            modalProps={{
              isBlocking: false,

            }}
            className={styles.dialog}
          >
            <div className={styles.contentClass}>
              {this.state.isRecordIsAlreadyInApproval
                ? <div style={{ marginBottom: '5px', color: 'red' }}>{this.state.resourceListItems["manage_my_tiles_record_already_in_approval_flow"]} </div>
                :
                <div>
                  <Icon iconName="Delete"></Icon>
                  <h2>{this.state.resourceListItems["remove_application_Dialog_Title"]} </h2>
                  <p>{this.state.resourceListItems["remove_application_Dialog_Text"]}</p>
                </div>
              }
            </div>
            <DialogFooter className={styles["ms-Dialog-actionsButton"]}>
              <DefaultButton text={this.state.resourceListItems["add_tile_cancel_text"]} onClick={() => this.setState({ hideDeleteDialog: true })} />

              {!this.state.isRecordIsAlreadyInApproval && (
                <PrimaryButton text={this.state.resourceListItems["remove_application_Dialog_Remove_Button_Text"]} onClick={this.yesDeleteItem.bind(this)} disabled={this.state.dissableRemoveButton} />
              )}

            </DialogFooter>
          </Dialog>
        </div>
      </React.Fragment >
    );
  }
  //capture color code
  private _onColRequestUndeletableChange = (
    event: React.MouseEvent<HTMLElement>,
    newText: boolean
  ) => {
    this.setState({ TileAvailableExternal: newText ? 0 : 1 });
    this.checkFormDataIsChanged(newText ? "0" : "1", "AvailableExternal");
  }

  private _onColorCodeChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
    this.selectedColorId = parseInt(item.key.toString());
    this.setState({
      tileRequestColorCode: this.selectedColorId,
      TileColorCodeId: this.selectedColorId,
    });
    this.checkFormDataIsChanged(item.key.toString(), "selectedColorId");
  }

  public async componentWillMount(): Promise<void> {

    //Color categories exist in state variable
    let colorCategory: any[] = this.state.lstColorMaster;
    if (colorCategory.length > 0) {
      colorCategory.map((v: IColorMasterList): void => {
        this.colorOptions.push({ key: v.Id, text: v.Title });
      });
    } else {
      //Color categories does not exist in state variable then query SPO list
      try {
        sp.web.lists.getByTitle(this.lstColorMaster).items.orderBy('Id', true).getAll().then((r: any[]): void => {
          if (r.length > 0) {
            r.map((v: IColorMasterList) => {
              this.colorOptions.push({ key: v.Id, text: v.Title });
            });
            this.setState({ lstColorMaster: r });
          }
        });
      } catch (error) {
        //console.log(error);
        this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "componentWillMount");

      }
    }
  }
  private ddItemsPerPageChanges(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) {

    this.setState({
      itemsPerPage: parseInt(option.key.toString())
    });

    let sliceItems = this.itemsForPagination.length > parseInt(option.key.toString()) ? parseInt(option.key.toString()) : this.itemsForPagination.length;
    let totalPage = this.itemsForPagination.length < parseInt(option.key.toString()) ? 1 : Math.ceil(this.itemsForPagination.length / parseInt(option.key.toString()));

    this.setState({
      currentPage: 1,
      myTilesItems: this.itemsForPagination.slice(0, sliceItems),
      totalPage: totalPage,
    });
  }
  private getTileRequestTitleErrorMessage = (value: string): string => {
    value = value.trim();
    return value == "" ? this.state.valueRequiredErrorMessage : value.length > 255 ? this.state.valueMaxCharacterErrorMessage : "";
  }
  private getTileRequestDescriptionErrorMessage = (value: string): string => {
    value = value.trim();
    return value == "" ? this.state.valueRequiredErrorMessage : value.length > 63999 ? this.state.valueMaxCharacterErrorMessage : "";
  }
  private getTileRequestKeywordsErrorMessage = (value: string): string => {
    value = value.trim();
    return value == "" ? this.state.valueRequiredErrorMessage : value.length > 63999 ? this.state.valueMaxCharacterErrorMessage : "";
  }
  private getTileRequestURLLinkErrorMessage = (value: string): string => {
    let errMsg = "";
    value = value.trim();
    if (value == "") {
      errMsg = this.state.resourceListItems["required_field_validation_message"];
    }
    else if (value.length > 255) {
      errMsg = this.state.valueMaxCharacterErrorMessage;
    }
    else if (!this.urlPattern.test(value)) {
      errMsg = this.state.resourceListItems["validation_invalid_url_text"];
    }
    else if (this.urlPattern.test(value)) {
      //console.log("Url is valid, now check for it alredy present or not");
      //await this.pnpHelper.checkApplicationUrlExistOrNot(this.state.currentItem["Id"], value).then(([result])  => {
      // console.log("Url exist ? ", urlFound);
      let urlFound = this.allApplicationTiles.some(tile => {
        if (tile.Link.toLowerCase() === value.toLowerCase()) {
          if (tile.IsActive) {
            errMsg = this.state.resourceListItems["validation_url_exist_text"];
            return errMsg;
          } else {
            errMsg = this.state.resourceListItems['validation_url_exist_disable_text'];
            return errMsg;
          }
        }
        else {
          errMsg = "";
        }
      });

      if (errMsg == "") {
        let urlFoundInApprovalRequest = this.allApplicationRequest.some(tile => {
          if (tile.Link.toLowerCase() === value.toLowerCase()) {
            errMsg = this.state.resourceListItems["add_tile_request_exist_msg"];
            return errMsg;
          }
          else {
            errMsg = "";
          }
        });
      }

      //});
    }
    //to disable Test URL

    return errMsg;

  }
  private getTileRequestOwnerEmailErrorMessage = (value: string): string => {
    value = value.trim();
    let errMsg =
      stringIsNullOrEmpty(value)
        ? this.state.resourceListItems['required_field_validation_message']
        : value.length > 500
          ? this.state.resourceListItems['validation_maximum_characters_text']
          : this.emailPattern.test(value) ? ""
            : this.state.resourceListItems['validation_invalid_email'];
    return errMsg;
  }
  private paginationChanged(page: any) {
    //console.log(this.itemsForPagination);
    this.setState({
      currentPage: page,
      myTilesItems: this.itemsForPagination.slice(this.state.itemsPerPage * page - this.state.itemsPerPage, this.state.itemsPerPage * page - this.state.itemsPerPage + this.state.itemsPerPage)
    });
  }


  public componentDidMount() {
    try {
      this.getMyAllTiles();
    } catch (ex) {

    }
  }

  private async getMyAllTiles(): Promise<void> {
    try {
      // let filterCondition = "OwnerEmail eq '" + this.props.wpContext.pageContext.user.loginName + "' and IsActive eq 1";
      // if (this.isOwnerMember) {
      //   filterCondition = "";
      // }
      Promise.all([
        this.pnpHelper.getMyTiles("")
      ]).then(([myTilesItems]) => {
        if (this.isOwnerMember) {
        } else {
          myTilesItems = myTilesItems.filter(v => v.OwnerEmail === this.props.wpContext.pageContext.user.loginName && v.IsActive === true);
        }
        myTilesItems = myTilesItems.sort((a, b) => { return b.ID - a.ID; });
        this.setState({
          myTilesItems: myTilesItems,
          isDataLoaded: true,
        });
        //console.log(myTilesItems)
        this._allItems = myTilesItems;
        this.itemsForPagination = myTilesItems;
        let sliceItems = myTilesItems.length > this.state.itemsPerPage ? this.state.itemsPerPage : myTilesItems.length;
        let totalPage = this._allItems.length < this.state.itemsPerPage ? 1 : Math.ceil(this._allItems.length / this.state.itemsPerPage);
        this.setState({
          currentPage: 1,
          totalPage: totalPage,
          myTilesItems: myTilesItems.slice(0, sliceItems),
        });
      });

    }//try end
    catch (error) {
      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "getMyAllTiles");
    }

  }
  private _onColumnClickSort = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {



    const { columns } = this.state;
    let { myTilesItems } = this.state;
    let isSortedDescending = column.isSortedDescending;

    // If we've sorted this column, flip it.
    if (column.isSorted) {
      isSortedDescending = !isSortedDescending;
    }

    // Sort the items.
    myTilesItems = this._copyAndSort(this._allItems, column.fieldName!, isSortedDescending);
    //console.log(myTilesItems);
    this._allItems = myTilesItems;
    this.itemsForPagination = myTilesItems;

    // Reset the items and columns to match the state.
    this.setState({
      myTilesItems: myTilesItems,
      columns: columns.map(col => {
        col.isSorted = col.key === column.key;

        if (col.isSorted) {
          col.isSortedDescending = isSortedDescending;
        }

        return col;
      }),
    });
    this.paginationChanged(1);
  }

  private _onFilter = (searchValue: string): void => {
    let filteredApplicationTileList: IManageMyTilesList[] = [];
    this._allItems.map((data: any, i: any) => {
      if (
        String(data["Title"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["SearchKeywords"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["IsActive"] == true ? "Active" : "Inactive").toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["Description"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1
      ) {
        filteredApplicationTileList.push(data);
      }
    });
    this.itemsForPagination = filteredApplicationTileList;
    let sliceItems = filteredApplicationTileList.length > this.state.itemsPerPage ? this.state.itemsPerPage : filteredApplicationTileList.length;
    //let totalPage = filteredApplicationTileList.length < this.state.itemsPerPage ? 1 : Math.round(this.itemsForPagination.length / this.state.itemsPerPage);
    let totalPage = this.itemsForPagination.length < this.state.itemsPerPage ? 1 : Math.ceil(this.itemsForPagination.length / this.state.itemsPerPage);
    this.setState({
      currentPage: 1,
      myTilesItems: filteredApplicationTileList.slice(0, sliceItems),
      totalPage: totalPage// Math.ceil(this.itemsForPagination.length / this.state.itemsPerPage),
    });
  }

  private hideAllPanel = (): void => {
    this.setState({
      showViewPanel: false,
      showEditPanel: false
    });
  }
  private showViewPanel = (item: any): void => {
    this.setState({
      status: "Item Loaded for View Tiles",
      currentItem: item,
      TileColorCategory: item['ColorCode'].Title == null ? "-" : item['ColorCode'].Title,
    });

    this.setState({ showViewPanel: true });
  }

  private showEditPanel = (item: any): void => {
    this.showError = false;
    this.pnpHelper.checkExitingRecordIsInApprovalInTrackRequest(item['Id']).then((itemFound: boolean) => {
      this.setState({
        isRecordIsAlreadyInApproval: itemFound,
        status: "Item Loaded for Edit Tiles",
        currentItem: item,
        TileId: item['Id'],
        TileTitle: item['Title'] == null ? "" : item['Title'],
        TileDescription: item['Description'] == null ? "" : item['Description'],
        TileKewords: item['SearchKeywords'] == null ? "" : item['SearchKeywords'],
        TileURLLink: item['Link'] == null ? "" : item['Link'],
        TileOwner: item['OwnerEmail'] == null ? "" : item['OwnerEmail'],
        TileApprovedDate: item['ApprovedDate'] == null ? "" : item['ApprovedDate'],
        TileColorCodeBgColor: item['ColorCode'].BgColor == null ? "" : item['ColorCode'],
        TileColorCodeForeColor: item['ColorCode'].ForeColor == null ? "" : item['ColorCode'].ForeColor,
        TileColorCodeId: item['ColorCode'].ID == 0 ? "" : item['ColorCode'].ID,
        TileColorCategory: item['ColorCode'].Title == null ? "-" : item['ColorCode'].Title,
        TileAvailableExternal: item['AvailableExternal'] == null ? 1 : item['AvailableExternal'],
        RequestedAction: 'Modification',
        formDataIsChanged: false,
        showEditPanel: true,
        valueValidEamilErrorMessage: "",
        valueValidUrlLinkErrorMessage: ""
      });
    });
    this._refreshValidationData(item['Id']);
  }
  private deleteTile(item: any) {

    this.pnpHelper.checkExitingRecordIsInApprovalInTrackRequest(item['Id']).then((itemFound: boolean) => {

      this.setState({
        status: "Item Loaded for delete Tiles",
        currentItem: item,
        TileId: item['Id'],
        TileTitle: item['Title'],
        TileDescription: item['Description'],
        TileKewords: item['SearchKeywords'],
        TileURLLink: item['Link'],
        TileOwner: item['OwnerEmail'],
        TileApprovedDate: item['ApprovedDate'],
        TileColorCodeBgColor: item['ColorCode'].BgColor,
        TileColorCodeForeColor: item['ColorCode'].ForeColor,
        TileColorCodeId: item['ColorCode'].ID,
        TileAvailableExternal: item['AvailableExternal'] == null ? 1 : item['AvailableExternal'],
        RequestedAction: 'Deletion',
        dissableRemoveButton: false,
        isRecordIsAlreadyInApproval: itemFound,
        hideDeleteDialog: false
      });
    });
  }

  private yesDeleteItem = (): void => {

    //Update the tile info in Application Master and set its status as ?
    //Create a new record in Application Request for Approval
    try {
      this.setState({ dissableRemoveButton: true });
      Promise.all([
        this.pnpHelper.createTileRequestToApplicationRequest(this.state)
      ]).then(([requestItem]) => {
        this.setState({ hideDeleteDialog: true });
        this.props.callBackForRequestSection(this.state.resourceListItems["add_tile_deletion_success_text"]);
        this.getMyAllTiles();
      });
    }
    catch (error) {
      //console.log(e);
      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "yesDeleteItem");
    }

  }

  private _onRenderTileEditFooterContent = (): JSX.Element => {
    return (
      <React.Fragment>
        <div>
          <div>
            {this.state.isRecordIsAlreadyInApproval
              ? <div style={{ marginBottom: '5px', color: 'red' }}>{this.state.resourceListItems["manage_my_tiles_record_already_in_approval_flow"]} </div>
              :
              <PrimaryButton
                onClick={this.validateTheForm.bind(this)}
                text={this.state.resourceListItems["add_tile_save_text"]}
                disabled={!this.state.formDataIsChanged}
              />
            }
            <DefaultButton
              onClick={this.hideAllPanel}
              text={this.state.resourceListItems["add_tile_cancel_text"]}
            />
          </div>
        </div>
      </React.Fragment>
    );
  }


  private _onTileTitleChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string): void => {
    this.setState(
      {
        TileTitle: newValue
      }
    );
    this.checkFormDataIsChanged(newValue, "TileTitle");
  }

  private _onTileDescriptionChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string): void => {
    this.setState(
      {
        TileDescription: newValue
      }
    );
    this.checkFormDataIsChanged(newValue, "TileDescription");
  }

  private _onTileKewordChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string): void => {
    this.setState(
      {
        TileKewords: newValue
      }
    );
    this.checkFormDataIsChanged(newValue, "TileKewords");
  }

  private _onTileLinkURLChange = async (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string): Promise<void> => {
    this.setState(
      {
        TileURLLink: newValue,
      }
    );
    this.checkFormDataIsChanged(newValue, "TileURLLink");

    //if url is valid then check it is already exist or not
    // if (this.urlPattern.test(newValue)) {
    //   await this.pnpHelper.checkApplicationUrlExistOrNot(this.state.currentItem["Id"], newValue).then((urlFound: boolean) => {
    //     this.setState({
    //       valueValidUrlLinkErrorMessage: urlFound ? this.state.resourceListItems["validation_url_exist_text"] : ""
    //     });
    //   });
    // }

  }


  private _onTileEmailChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue: string): void => {

    this.setState(
      {
        TileOwner: newValue,
      }
    );
    this.checkFormDataIsChanged(newValue, "TileOwner");
  }

  private validateTheForm = async () => {
    try {
      this.ctrTitle.current.focus();
      this.ctrDescription.current.focus();
      this.ctrKeywords.current.focus();
      this.ctrURL.current.focus();
      if (this.isOwnerMember) {
        this.ctrOwnerEmail.current.focus();
      }

      this.showError = this.selectedColorId === 0 ? true : false;
      if (
        this.state.valueValidUrlLinkErrorMessage != "" ||
        this.state.TileTitle.trim() === "" ||
        this.state.TileTitle.length > 255 ||
        this.state.TileDescription.trim() === "" ||
        this.state.TileDescription.length > 63999 ||
        this.state.TileKewords.trim() === "" ||
        this.state.TileKewords.length > 63999 ||
        this.urlPattern.test(this.state.TileURLLink) == false ||
        this.state.TileURLLink.length > 255 ||
        this.emailPattern.test(this.state.TileOwner) == false ||
        this.state.TileOwner.length > 255 ||
        this.selectedColorId == 0 ||
        this.getTileRequestURLLinkErrorMessage(this.state.TileURLLink) !== ""
      ) {
        //alert("Invalid Form")
        this.setState({
          showEditPanel: true,
        });
      }
      // checking email and url type
      else {
        //alert("Valida Form")
        this.setState({
          formDataIsChanged: false, //disabling save items 
        });

        try {
          Promise.all([
            this.pnpHelper.createTileRequestToApplicationRequest(this.state)
          ]).then(([requestItem]) => {
            this.setState({
              showEditPanel: false,
            });
            this.props.callBackForRequestSection(this.state.resourceListItems["add_tile_modification_success_text"]);
          });
        } catch (error) {
          this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "Create tile request");
        }



      }
    } catch (error) {
      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "validateTheForm");
    }
  }

  private checkFormDataIsChanged(newValue: string, ControllName: string) {

    setTimeout(() => {
      if (this.state.showEditPanel == true) {
        if (
          this.state.currentItem['Title'] === this.state.TileTitle.trim() &&
          this.state.currentItem['Description'] === this.state.TileDescription.trim() &&
          this.state.currentItem['SearchKeywords'] === this.state.TileKewords.trim() &&
          this.state.currentItem['Link'] === this.state.TileURLLink.trim() &&
          this.state.currentItem['OwnerEmail'] === this.state.TileOwner.trim() &&
          this.state.currentItem['ColorCode'].ID === this.state.TileColorCodeId &&
          this.state.currentItem['AvailableExternal'] === this.state.TileAvailableExternal
        ) {
          this.setState({ formDataIsChanged: false });
        } else {
          this.setState({ formDataIsChanged: true });
        }
      }
    }, 1000);


  }


  private _mathRound(num: any, prec: any) {
    var magn = Math.pow(10, prec);
    return Math.ceil(num * magn) / magn;
  }

  private _copyAndSort<T>(items: any[], columnKey: string, isSortedDescending?: boolean): any[] {
    const key = columnKey as keyof any;

    //return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    return items.slice(0).sort((a: any, b: any) => ((isSortedDescending ? a[key].toString().toUpperCase() < b[key].toString().toUpperCase() : a[key].toString().toUpperCase() > b[key].toString().toUpperCase()) ? 1 : -1));

  }
  private _refreshValidationData = async (itemID: number) => {
    //populate application master data
    try {
      sp.web.lists.getByTitle(this.ApplicationMasterListName).items
        .select('Id', 'Link', 'IsActive')
        .filter("Id ne " + itemID)
        .top(5000)
        .get()
        .then((applicationTiles: any[]) => {
          this.allApplicationTiles = applicationTiles;
        });

      //get Tile request data for validaion
      sp.web.lists.getByTitle(this.lstTileRequest)
        .items.select('Id', 'Link')
        .filter(`ApprovalStatus eq 'Waiting for Approval'`)
        .top(5000)
        .get()
        .then((r): void => {
          //if (r.length > 0) {
            this.allApplicationRequest = r;
         // }
        });
    } catch (error) {
      //console.log(e);
      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "_refreshValidationData");
    }

  }

}






