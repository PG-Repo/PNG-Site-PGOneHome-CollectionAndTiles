import * as React from 'react';
import { Text, Dropdown, IDropdownStyles, PrimaryButton, IDropdownOption, Link, IColumn, Icon, SearchBox, DetailsList, SelectionMode, DetailsListLayoutMode, Shimmer, Panel, PanelType, Label } from 'office-ui-fabric-react';
import { PnPHelper } from '../PnPHelper/PnPHelper';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import styles from '../TrackRequest/TrackRequest.module.scss';
import { SPPermission } from '@microsoft/sp-page-context';
import Pagination from 'office-ui-fabric-react-pagination';
interface ITrackRequestList {
  Title: string;
  ID?: string;
  Link: string;
}
interface ITrackRequestState {
  status: string;
  resourceListItems: any[];
  ddRequestType: string;
  ddRequestStatus: string;
  isDataLoaded: boolean;
  columns: IColumn[];

  trackRequestItems: ITrackRequestList[];
  currentItem: ITrackRequestList[];
  trackRequestCollectionApplication: ITrackRequestList[];
  TileColorCategory: string;

  showApplicationPanel: boolean;
  showCollectionPanel: boolean;
  showTileSection:boolean;


  currentPage?: number;
  totalPage?: number;
  itemsPerPage?: number;

  currentRequestType: string;

}
interface ITrackRequestProps {
  wpContext: WebPartContext;
  resourceListItems: any[];
  callBackForRequestSection: any;

}

const ddTrackRequestTypeOptions = [
  { key: 'Tile Request', text: 'Tile Request', listName: "ApplicationRequestsTemp" },
  { key: 'Collection Request', text: 'Collection Request', listName: "CollectionRequests" },
];
const ddTrackRequestStatusOptions = [
  { key: 'Waiting for Approval', text: 'Waiting for Approval' },
  { key: 'Approved', text: 'Approved' },
  { key: 'Rejected', text: 'Rejected' },
];
const ddNoOFItems = [
  { key: 10, text: '10' },
  { key: 20, text: '20' },
  { key: 30, text: '30' },
  { key: 50, text: '50' },
  { key: 100, text: '100' },
];
const ddOptionsStyles: Partial<IDropdownStyles> = { dropdown: { width: 100, height: 28, marginLeft: 6, marginRight: 6, marginTop: -7, marginBottom: -7 } };


const ddTrackRequestStatusOptionsStyles: Partial<IDropdownStyles> = {};
const iconProps = { iconName: 'Calendar' };
export class TrackRequest extends React.Component<ITrackRequestProps, ITrackRequestState> {

  private pnpHelper: PnPHelper;
  private _allItems: ITrackRequestList[];
  private itemsForPagination: ITrackRequestList[];

  private errTitle: string = "PNG-Site-PGOneHome-CollectionAndTiles";
  private errModule: string = "TrackRequests.tsx";

  private isOwnerMember = this.props.wpContext.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
  constructor(props: ITrackRequestProps, state: ITrackRequestState) {
    super(props);
    const columns: IColumn[] = [
      {
        key: 'ID', name: 'Sr. No', fieldName: 'ID', minWidth: 20, maxWidth: 60, isResizable: true, isCollapsible: true,
        onRender: (item, index) => (
          <div>
            {(this.state.itemsPerPage * this.state.currentPage + index + 1) - (this.state.itemsPerPage)}
          </div>
        ),
        //onColumnClick:this._onColumnClickSort.bind(this),
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
        key: 'RequestedAction',
        name: 'Requested Action',
        ariaLabel: 'Requested Action',
        isSorted: false,
        sortAscendingAriaLabel: 'Sort A to Z',
        sortDescendingAriaLabel: 'Sort Z to A',
        onColumnClick: this._onColumnClickSort.bind(this),
        isRowHeader: true,
        data: 'string',
        isPadded: true,
        fieldName: 'RequestedAction',
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        isCollapsible: true,
      },
      {
        key: 'RequestType', name: 'Request Type', minWidth: 140, maxWidth: 150, isResizable: true, isCollapsible: true,
        onRender: (item) => (
          <div style={{ color: "#323130", fontSize: 14 }}>
            {this.state.ddRequestType === 'Tile Request' ? 'Tile' : 'Collection'}
          </div>)
      },
      {
        key: 'Status',
        name: 'Status',
        fieldName: 'ApprovalStatus',
        minWidth: 140,
        maxWidth: 150,
        isResizable: true,
        isCollapsible: true,
        onRender: (item) => (
          <div style={{ color: "#323130", fontSize: 14 }}>
            {item.ApprovalStatus}
          </div>)

      },
      // {
      //   key: 'RequestedDate', name: 'Requested Date', fieldName: 'RequestedDate', minWidth: 125, maxWidth: 125, isResizable: true, isCollapsible: true,
      //   onRender: (item) => (
      //     <div>
      //       {new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "2-digit", hour: 'numeric', minute: 'numeric', }).format(new Date(item.RequestedDate))}
      //     </div>),
      //   data: 'Date',
      //   //onColumnClick: this._onColumnClickSort.bind(this),  onClick={() => this.showTileViewPanel(item)}onClick={() => this.showCollectionViewPanel(item)}
      // },

      {
        key: 'View', name: 'Action', fieldName: '', minWidth: 50, maxWidth: 50, isResizable: true, isCollapsible: false,
        onRender: (item) => (
          <div>
            {this.state.currentRequestType == "Tile Request"
              ? <Link aria-label="View" title="View" onClick={() => this.showTileViewPanel(item)}><Icon className="icons" iconName="View" /></Link>
              : <Link aria-label="View" title="View" onClick={() => this.showCollectionViewPanel(item)}><Icon className="icons" iconName="View" /></Link>
            }


          </div>
        ),

      }
    ];
    this.state = {
      status: "Ok",
      resourceListItems: this.props.resourceListItems,
      ddRequestType: "Tile Request",
      currentRequestType: "Tile Request",
      ddRequestStatus: "Waiting for Approval",
      columns: columns,
      isDataLoaded: false,
      trackRequestItems: [],
      currentItem: [],
      trackRequestCollectionApplication: [],

      showApplicationPanel: false,
      showCollectionPanel: false,
      showTileSection:true,
      TileColorCategory: "",

      currentPage: 1,
      totalPage: 1,
      itemsPerPage: 10,
    };
    this.pnpHelper = new PnPHelper(this.props.wpContext);
  }
  public render(): React.ReactElement<ITrackRequestProps> {
    return (
      <div className={styles.trackRequest}>
        <div className={`ms-Grid`}>
          <div className="ms-Grid-row ">
            <div className="ms-Grid-col ms-sm10 ms-md12 ms-lg12 ms-xl12">
              <h2 className="ms-font-m ms-fontWeight-semibold">{this.state.resourceListItems['track_request_label']}</h2>
            </div>
          </div>
          <hr role="presentation"></hr>
          <div className={`ms-Grid-row ${styles.gridRowFilter}`} >
            <div className="ms-Grid-col ms-sm5 ms-md4 ms-lg4  ms-xl2">
              <label id="ddlRequestType" className="ms-font-m ms-fontWeight-semibold">{this.state.resourceListItems['review_request_type_label']}</label>
            </div>
            <div className="ms-Grid-col ms-sm7 ms-md8 ms-lg5 ms-xl3">
              <Dropdown
                // placeholder="Select an option"
                label=""
                options={ddTrackRequestTypeOptions}
                styles={ddTrackRequestStatusOptionsStyles}
                selectedKey={this.state.ddRequestType}
                onChange={this.ddRequestTypeChanges.bind(this)}
                aria-labelledby="ddlRequestType"
               
              />
            </div>

            <div className="ms-Grid-col ms-sm5 ms-md4 ms-lg4  ms-xl2">
              <label id="ddlRequestStatus" className="ms-font-m ms-fontWeight-semibold">{this.state.resourceListItems['review_request_status_label']}</label>
            </div>
            <div className="ms-Grid-col ms-sm7 ms-md8 ms-lg5 ms-xl3">
              <Dropdown
                // placeholder="Select an option"
                label=""
                options={ddTrackRequestStatusOptions}
                styles={ddTrackRequestStatusOptionsStyles}
                selectedKey={this.state.ddRequestStatus}
                onChange={this.ddRequestStatusChanges.bind(this)}
                aria-labelledby="ddlRequestStatus"
                
              />
            </div>

            <div className="ms-Grid-col ms-sm2 ms-md2 ms-lg2 ms-xl2">
              <PrimaryButton text={this.state.resourceListItems['review_submit_text']} onClick={() => this.getTrckRequest()} />
            </div>
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
                onChange={this.ddItemsPerPageChanges.bind(this)}
                aria-labelledby="lbl1 lbl2"
                    
              />
                  <label id="lbl2">entries</label>
                  </div>
          </div>
          <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg3 ms-xl3 searchBoxPG detailsListContoll">
            <SearchBox placeholder={this.state.resourceListItems["review_search_placeholder_text"]} onSearch={this._onFilterApplication.bind(this)} onChange={this._onFilterApplication.bind(this)} />
          </div>
        </div>
        <div className={`ms-Grid-row ${styles.gridRowDetailsList}`}>
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12" >
            <DetailsList items={this.state.trackRequestItems}
              setKey="none"
              ariaLabelForListHeader="Column headers. Click to sort."
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              ariaLabelForSelectionColumn="Toggle selection"
              checkButtonAriaLabel="Row checkbox"
              columns={this.state.columns}
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              selectionMode={SelectionMode.none}
              className={styles.TrackRequestItems}

            />

            <Shimmer className={styles.shimmerLoader} isDataLoaded={this.state.isDataLoaded} />
            <Shimmer className={styles.shimmerLoader} width="75%" isDataLoaded={this.state.isDataLoaded} />
            <Shimmer className={styles.shimmerLoader} width="50%" isDataLoaded={this.state.isDataLoaded} />

            {this.state.trackRequestItems.length == 0 && this.state.isDataLoaded && (
              <div> No records found.</div>
            )}
            {this.state.trackRequestItems.length > 0 && this.state.isDataLoaded && (
              <div className="paginationList">
                <Pagination currentPage={this.state.currentPage} totalPages={this.state.totalPage} onChange={(page: any) => { this.paginationChanged(page); }} />
              </div>
            )
            }
          </div>


        </div>

        {/* View Tiles Panel */}
        <Panel
          isLightDismiss
          className={styles.ViewTilesPanel}
          headerText={this.state.resourceListItems['track_request_label']}
          isOpen={this.state.showApplicationPanel}
          onDismiss={this.hideAllPanel}
          closeButtonAriaLabel="Close"
          type={PanelType.smallFixedFar}
          style={{ fontWeight: "bold" }}
          // onRenderFooterContent={this._onRenderTileEditFooterContent}
          isFooterAtBottom={true}
          headerClassName={styles.headerText}
        >
          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_request_type_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentRequestType}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_requested_action_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['RequestedAction'] == undefined ? "-" : this.state.currentItem['RequestedAction']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['add_tile_title_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['Title'] == undefined ? "-" : this.state.currentItem['Title']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['add_tile_description_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['Description'] == undefined ? "-" : this.state.currentItem['Description']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['add_tile_keywords_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['SearchKeywords'] == undefined ? "-" : this.state.currentItem['SearchKeywords']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['add_tile_url_link_label']}</Label>
          <Label className={styles["ms-font-m"]} >
            {this.state.currentItem['Link']}<br></br>
            <Link data-interception="off" href={this.state.currentItem['Link'] == undefined ? "#" : this.state.currentItem['Link']} target="_blank">{this.state.resourceListItems['add_tile_try_link_label']}</Link>
          </Label>


          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['add_tile_owneremail_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['OwnerEmail'] == undefined ? "-" : this.state.currentItem['OwnerEmail']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_category_label"]}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.TileColorCategory == undefined ? "-" : this.state.TileColorCategory}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems["add_tile_available_external_label"]}</Label>
            <Label className={styles["ms-font-m"]}>
              {this.state.currentItem['AvailableExternal'] == null ? "-" : this.state.currentItem['AvailableExternal'] == 0 ? "Yes" : "No"}
            </Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_requested_date_label']}</Label>
          <Label className={styles["ms-font-m"]}>
            {this.state.currentItem['RequestedDate'] == undefined ? "-" :
              //new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "2-digit",hour: 'numeric', minute: 'numeric', }).format(new Date(this.state.currentItem['RequestedDate']))}
              new Date(this.state.currentItem['RequestedDate'].toString()).toLocaleString()}
          </Label>
          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['approval_status_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['ApprovalStatus'] == undefined ? "-" : this.state.currentItem['ApprovalStatus']}</Label>
          
          {this.state.currentItem['ApprovalStatus']!="Waiting for Approval" &&(
            <div>
          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_decision_date_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['DecisionDate'] == undefined ? "-" : new Date(this.state.currentItem['DecisionDate'].toString()).toLocaleString()}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_decision_by_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['DecisionBy'] == undefined ? "-" : this.state.currentItem['DecisionBy']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_decision_comments_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['DecisionComments'] == undefined ? "-" : this.state.currentItem['DecisionComments']}</Label>
 
            </div>
          )}
          
        </Panel>


        {/* View Collection Panel */}
        <Panel
          isLightDismiss
          className={styles.ViewTilesPanel}
          headerText={this.state.resourceListItems['track_request_label']}
          isOpen={this.state.showCollectionPanel}
          onDismiss={this.hideAllPanel}
          closeButtonAriaLabel="Close"
          type={PanelType.smallFixedFar}
          style={{ fontWeight: "bold" }}
          // onRenderFooterContent={this._onRenderTileEditFooterContent}
          isFooterAtBottom={true}
          headerClassName={styles.headerText}
        >
          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_collection_type_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentRequestType}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_requested_action_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['RequestedAction'] == undefined ? "-" : this.state.currentItem['RequestedAction']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems["new_collection_name"]}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['Title'] == undefined ? "-" : this.state.currentItem['Title']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems["new_collection_make_public"]}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['PublicCollection'] == undefined ? "-" : this.state.currentItem['PublicCollection'] == 1 ? "Yes" : "No"}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['new_collection_description']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['Description'] == undefined ? "-" : this.state.currentItem['Description']}</Label>

           
            <Label className={styles["ms-Label"]}>{this.state.resourceListItems['tiles']}</Label>
            <Label className={`${styles.TileCollections} ${styles["ms-font-m"]}`}>
              {this.state.trackRequestCollectionApplication.length > 0 && (
                this.state.trackRequestCollectionApplication.map(applicationItems => (
                  <a href={applicationItems.Link} target="_blank" data-interception="off" title={applicationItems.Title}>{applicationItems.Title}, </a>
                ))
              )}
          </Label>
          
          

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_collection_owner_label']}</Label>
          <Label className={styles["ms-font-m"]}>
            {this.state.currentItem['CollectionOwner'] == undefined ? "-" : this.state.currentItem['CollectionOwner'].Email}
          </Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_requested_date_label']}</Label>
          <Label className={styles["ms-font-m"]}>
            {this.state.currentItem['RequestedDate'] == null ? "-" :
              new Date(this.state.currentItem['RequestedDate'].toString()).toLocaleString()
              //new Intl.DateTimeFormat("en-US", { year: "numeric", month: "long", day: "2-digit",hour: 'numeric', minute: 'numeric', }).format(new Date(this.state.currentItem['RequestedDate']))
            }
          </Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['approval_status_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['ApprovalStatus'] == undefined ? "-" : this.state.currentItem['ApprovalStatus']}</Label>
          
          {this.state.currentItem['ApprovalStatus']!="Waiting for Approval" &&(
            <div>
          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_decision_date_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['DecisionDate'] == null ? "-" : new Date(this.state.currentItem['DecisionDate'].toString()).toLocaleString()}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_decision_by_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['DecisionBy'] == null ? "-" : this.state.currentItem['DecisionBy']}</Label>

          <Label className={styles["ms-Label"]}>{this.state.resourceListItems['review_decision_comments_label']}</Label>
          <Label className={styles["ms-font-m"]}>{this.state.currentItem['DecisionComments'] == null ? "-" : this.state.currentItem['DecisionComments']}</Label>
 
            </div>
          )}

        </Panel>
      </div>
    );
  }

  public componentDidMount() {
    try {
      this.getTrckRequest();
    } catch (error) {

      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "componentDidMount");
    }
  }

  private async getAllRequestForApplication(): Promise<void> {
    try {
      //let currentUserTNumber = await this.pnpHelper.userProps("TNumber");
      //let currentUserItem = await this.pnpHelper.getCurrentUserItemID(currentUserTNumber);
      //let filterCondition = "(OwnerEmail eq '" + this.props.wpContext.pageContext.user.loginName + "' or RequestedById eq '" + currentUserItem[0].ID + "' and ApprovalStatus eq '" + this.state.ddRequestStatus + "')";

      // let filterCondition = "(OwnerEmail eq '" + this.props.wpContext.pageContext.user.loginName + "' and ApprovalStatus eq '" + this.state.ddRequestStatus + "')";
      // if (this.isOwnerMember) {
      //   filterCondition = "(ApprovalStatus eq '" + this.state.ddRequestStatus + "')";
      // }
      Promise.all([
        this.pnpHelper.getTrackRequestForApplication("")
      ]).then(([trackRequestItems]) => {

        //Fitering items
        if (this.isOwnerMember) {
          trackRequestItems = trackRequestItems.filter(v => v.ApprovalStatus === this.state.ddRequestStatus);
        } else {
          trackRequestItems = trackRequestItems.filter(v => v.OwnerEmail === this.props.wpContext.pageContext.user.loginName && v.ApprovalStatus === this.state.ddRequestStatus);
        }
        trackRequestItems = trackRequestItems.sort((a, b) => { return b.ID - a.ID; });
        this.setState({
          trackRequestItems: trackRequestItems,
          isDataLoaded: true,
        });
        // console.log(trackRequestItems)
        this._allItems = trackRequestItems;
        this.itemsForPagination = trackRequestItems;

        let sliceItems = trackRequestItems.length > this.state.itemsPerPage ? this.state.itemsPerPage : trackRequestItems.length;
        let totalPage = this._allItems.length < this.state.itemsPerPage ? 1 : Math.ceil(this._allItems.length / this.state.itemsPerPage);
        this.setState({
          currentPage: 1,
          totalPage: totalPage,
          trackRequestItems: trackRequestItems.slice(0, sliceItems),
        });


      });

    }//try end
    catch (error) {
      //console.log(e);
      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "getAllRequestForApplication");
    }

  }

  private async getAllRequestForCollection(): Promise<void> {

    try {
      let currentUserTNumber = await this.pnpHelper.userProps("TNumber");
      let currentUserItem = await this.pnpHelper.getCurrentUserItemID(currentUserTNumber);

      // let filterCondition = "(CollectionOwnerId eq '" + currentUserItem[0].ID + "' and ApprovalStatus eq '" + this.state.ddRequestStatus + "')";
      // if (this.isOwnerMember) {
      //   filterCondition = "(ApprovalStatus eq '" + this.state.ddRequestStatus + "')";
      // }
      Promise.all([
        this.pnpHelper.getTrackRequestForCollection("")
      ]).then(([trackRequestItems]) => {
        //Fitering items
        if (this.isOwnerMember) {
          trackRequestItems = trackRequestItems.filter(v => v.ApprovalStatus === this.state.ddRequestStatus);
        } else {
          trackRequestItems = trackRequestItems.filter(v => v.CollectionOwner.Id === currentUserItem[0].ID && v.ApprovalStatus === this.state.ddRequestStatus);
        }
        trackRequestItems = trackRequestItems.sort((a, b) => { return b.ID - a.ID; });

        this.setState({
          trackRequestItems: trackRequestItems,
          isDataLoaded: true,

        });
        // console.log(trackRequestItems)
        this._allItems = trackRequestItems;
        this.itemsForPagination = trackRequestItems;

        let sliceItems = trackRequestItems.length > this.state.itemsPerPage ? this.state.itemsPerPage : trackRequestItems.length;
        let totalPage = this._allItems.length < this.state.itemsPerPage ? 1 : Math.ceil(this._allItems.length / this.state.itemsPerPage);
        this.setState({
          currentPage: 1,
          totalPage: totalPage,
          trackRequestItems: trackRequestItems.slice(0, sliceItems),
        });

      });

    }//try end
    catch (error) {
      //console.log(e);
      this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "getAllRequestForCollection");
    }

  }


  private ddRequestStatusChanges(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) {
    //console.log(option.key);
    this.setState({
      ddRequestStatus: option.key.toString()
    });
  }
  private ddRequestTypeChanges(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) {
    //console.log(option.key);
    this.setState({
      ddRequestType: option.key.toString()
    });
  }
  private _onFilterApplication = (searchValue: string): void => {
    let filteredApplicationTileList: ITrackRequestList[] = [];
    this._allItems.map((data: any, i: any) => {
      if (
        String(data["Title"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["SearchKeywords"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["IsActive"] == true ? "Active" : "Inactive").toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["ID"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
        String(data["Description"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1
      ) {
        filteredApplicationTileList.push(data);
      }
    });
    this.itemsForPagination = filteredApplicationTileList;
    let sliceItems = filteredApplicationTileList.length > this.state.itemsPerPage ? this.state.itemsPerPage : filteredApplicationTileList.length;
    let totalPage = this.itemsForPagination.length < this.state.itemsPerPage ? 1 : Math.ceil(this.itemsForPagination.length / this.state.itemsPerPage);
    this.setState({
      currentPage: 1,
      trackRequestItems: filteredApplicationTileList.slice(0, sliceItems),
      totalPage: totalPage,
    });
  }




  private getTrckRequest(): void {

    this.setState({
      trackRequestItems: [],
      isDataLoaded: false,
      currentRequestType: this.state.ddRequestType

    });

    if (this.state.ddRequestType == "Tile Request") {
      this.getAllRequestForApplication();
    } else {
      this.getAllRequestForCollection();
    }

  }
  private hideAllPanel = (): void => {
    this.setState({
      showApplicationPanel: false,
      showCollectionPanel: false,
    });
  }
  private showTileViewPanel = (item: any): void => {
    //console.log("Tile Request",item);
    this.setState({
      status: "Item Loaded for View Tiles",
      currentItem: item,
      showApplicationPanel: true,
      showCollectionPanel: false,
      TileColorCategory: item['ColorCode'].Title == null ? "-" : item['ColorCode'].Title,
    });
  }
  private showCollectionViewPanel = (item: any): void => {

    this.setState({
      trackRequestCollectionApplication: []
    });
    //console.log("Collection Request",item);
    this.setState({
      status: "Item Loaded for View Tiles",
      currentItem: item,
      showApplicationPanel: false,
      showCollectionPanel: true,
      //showTileSection:item.RequestedAction="Addition"?true:false
    });

    if (this.state.ddRequestType == "Collection Request") {
      if(item.RequestedAction=="Addition"){
        try {
          //Get recored from the CollectionApplicationMatrixRequests for current Collection request
          Promise.all([
            this.pnpHelper.getCurrentCollectionApplication(item.ID)
          ]).then(([CollectionApplicationMatrixRequests]) => {
  
            //console.log("Collection Request Matrix",CollectionApplicationMatrixRequests)
  
            //Get the application information from the application master based on the ids presented in CollectionApplicationMatrixRequests
            if(CollectionApplicationMatrixRequests.length>0){
              let filterCodition = "";
              {
                CollectionApplicationMatrixRequests.map(items => (
                  //console.log(items.ApplicationIDId);
                  filterCodition += "(Id eq " + items.ApplicationIDId + ") or "
                ));
              }
              if (filterCodition.length > 0) { filterCodition = filterCodition.slice(0, filterCodition.length - 4); }
              // console.log("get application from the application master",filterCodition);
    
              Promise.all([
                this.pnpHelper.getMyTilesByItsId(filterCodition)
              ]).then(([applicationMasterItems]) => {
    
                // console.log(applicationMasterItems)
                this.setState({
                  trackRequestCollectionApplication: applicationMasterItems
                });
              });
            }
  
  
  
          });
  
        }//try end
        catch (e) {
          //console.log(e);
          this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", e, "showCollectionViewPanel");
        }
      }
      else{
        try {
          //Get recored from the CollectionApplicationMatrix for current Collection request
          Promise.all([
            this.pnpHelper.getCurrentCollectionApplicationFromCollectionApplicationMatrx(item.ExistingItemID)
          ]).then(([CollectionApplicationMatrixRequests]) => {
  
            //console.log("Collection Request Matrix",CollectionApplicationMatrixRequests)
  
            //Get the application information from the application master based on the ids presented in CollectionApplicationMatrixRequests
            if(CollectionApplicationMatrixRequests.length>0){
              let filterCodition = "";
              {
                CollectionApplicationMatrixRequests.map(items => (
                  //console.log(items.ApplicationIDId);
                  filterCodition += "(Id eq " + items.ApplicationIDId + ") or "
                ));
              }
              if (filterCodition.length > 0) { filterCodition = filterCodition.slice(0, filterCodition.length - 4); }
              // console.log("get application from the application master",filterCodition);
    
              Promise.all([
                this.pnpHelper.getMyTilesByItsId(filterCodition)
              ]).then(([applicationMasterItems]) => {
    
                // console.log(applicationMasterItems)
                this.setState({
                  trackRequestCollectionApplication: applicationMasterItems
                });
              });
            }
  
  
  
          });
  
        }//try end
        catch (e) {
          //console.log(e);
          this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", e, "showCollectionViewPanel");
        }
      }
      
    }

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
  private ddItemsPerPageChanges(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) {

    this.setState({
      itemsPerPage: parseInt(option.key.toString())
    });

    let sliceItems = this.itemsForPagination.length > parseInt(option.key.toString()) ? parseInt(option.key.toString()) : this.itemsForPagination.length;
    //let totalPage = this.itemsForPagination.length < parseInt(option.key.toString()) ? 1 : this._mathRound(this.itemsForPagination.length / parseInt(option.key.toString()), 0);
    let totalPage = this.itemsForPagination.length < parseInt(option.key.toString()) ? 1 : Math.ceil(this.itemsForPagination.length / parseInt(option.key.toString()));

    this.setState({
      currentPage: 1,
      trackRequestItems: this.itemsForPagination.slice(0, sliceItems),
      totalPage: totalPage,
    });
  }
  private paginationChanged(page: any) {
    //console.log(this.itemsForPagination);
    //console.log(this.state.itemsPerPage * page - this.state.itemsPerPage);
    //console.log(this.state.itemsPerPage * page - this.state.itemsPerPage + this.state.itemsPerPage);
    this.setState({
      currentPage: page,
      trackRequestItems: this.itemsForPagination.slice(this.state.itemsPerPage * page - this.state.itemsPerPage, this.state.itemsPerPage * page - this.state.itemsPerPage + this.state.itemsPerPage)
    });
  }
  private _onColumnClickSort = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {



    const { columns } = this.state;
    let { trackRequestItems } = this.state;
    let isSortedDescending = column.isSortedDescending;

    // If we've sorted this column, flip it.
    if (column.isSorted) {
      isSortedDescending = !isSortedDescending;
    }

    // Sort the items.
    trackRequestItems = this._copyAndSort(this._allItems, column.fieldName!, isSortedDescending);
    //console.log(myTilesItems);
    this._allItems = trackRequestItems;
    this.itemsForPagination = trackRequestItems;

    // Reset the items and columns to match the state.
    this.setState({
      trackRequestItems: trackRequestItems,
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

}