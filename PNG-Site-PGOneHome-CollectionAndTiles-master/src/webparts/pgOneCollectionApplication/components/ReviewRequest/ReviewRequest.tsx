import * as React from 'react';
import { IReviewRequestProps, IReviewRequestState } from './IReviewRequest';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import styles from "./ReviewRequest.module.scss";
import { PnPHelper } from '../PnPHelper/PnPHelper';
import { ErrorLogging } from "../ErrorLogging/ErrorLogging";
import { Dropdown, SearchBox, DetailsList, PrimaryButton, IDropdownOption, IColumn, DetailsListLayoutMode, Link, Icon, SelectionMode, Panel, PanelType, TextField, MessageBar, MessageBarType, DefaultButton, Label, ITextField, Shimmer, Sticky, StickyPositionType, Toggle, Spinner, ScrollablePane, ScrollbarVisibility, TooltipHost, IRenderFunction, IDetailsHeaderProps, ConstrainMode, IDetailsFooterProps, DetailsRow, IDetailsRow, IDetailsRowProps, ShimmeredDetailsList, IDropdownStyles } from 'office-ui-fabric-react';
import { ITileRequest } from '../Common/ITileRequest';
import { ICollectionRequest } from '../Common/ICollectionRequest';
import { IColorMasterList } from '../Common/IColormasterList';
import { stringIsNullOrEmpty } from '@pnp/common';
import Pagination from 'office-ui-fabric-react-pagination';
import { ICorpCollectionQueue } from '../Common/ICorpCollectionQueue';

export class ReviewRequest extends React.Component<IReviewRequestProps, IReviewRequestState> {
    private pnpHelper: PnPHelper;
    private errorLogging: ErrorLogging;

    private detailedListColumns: IColumn[];
    private searchData: any[];
    private lstTileMaster = "ApplicationMaster";
    private lstTileRequest = "ApplicationRequests";
    private lstColAppMatrix = "CollectionApplicationMatrix";
    private lstCollectionMaster = "CollectionMaster";
    private lstCollectionRequest = "CollectionRequests";
    private lstColAppMatrixRequests = "CollectionApplicationMatrixRequests";
    private lstColorMaster = "ColorMaster";
    private lstUserMaster = "UserMaster";
    private lstCorpCollectionQueue = "CorpCollectionQueue";

    private colorOptions: IDropdownOption[] = [{ key: 0, text: 'Select a category' }];
    private currentUserTNumber: string = "";
    private currentUserEmail: string = "";
    private selectedTiles: any[] = [];
    private selectedColorId: number = 1;
    private showError: boolean = false;
    private searchText: string = "";
    private currentRequestType: string = 'Tile Request';
    private currentRequestStatus: string = 'Waiting for Approval';
    private isDataRefresh: boolean = false;

    private errTitle: string = "PNG-Site-PGOneHome-CollectionAndTiles";
    private errModule: string = "ReviewRequest.tsx";

    private ctrTitle: React.RefObject<ITextField>;
    private ctrDescription: React.RefObject<ITextField>;
    private ctrKeywords: React.RefObject<ITextField>;
    private ctrURL: React.RefObject<ITextField>;
    private ctrOwnerEmail: React.RefObject<ITextField>;
    private ctrDecisionComments: React.RefObject<ITextField>;

    private ctrColTitle: React.RefObject<ITextField>;
    private ctrColDescription: React.RefObject<ITextField>;
    private ctrColDecisionComments: React.RefObject<ITextField>;

    private ddNoOFItems = [
        { key: 10, text: '10' },
        { key: 20, text: '20' },
        { key: 30, text: '30' },
        { key: 50, text: '50' },
        { key: 100, text: '100' },
    ];
    private ddOptionsStyles: Partial<IDropdownStyles> = {
        dropdown: {
            width: 70,
            height: 28,
            marginLeft: 6,
            marginRight: 6,
            marginTop: -7,
            marginBottom: -7,
        }
    };


    constructor(props: IReviewRequestProps) {
        super(props);
        this.pnpHelper = new PnPHelper(this.props.context);
        this.errorLogging = new ErrorLogging(this.props.context);

        this.ctrTitle = React.createRef();
        this.ctrDescription = React.createRef();
        this.ctrKeywords = React.createRef();
        this.ctrURL = React.createRef();
        this.ctrOwnerEmail = React.createRef();
        this.ctrDecisionComments = React.createRef();

        this.ctrColTitle = React.createRef();
        this.ctrColDescription = React.createRef();
        this.ctrColDecisionComments = React.createRef();

        this.detailedListColumns = [
            {
                fieldName: 'Id',
                name: 'Sr. No',
                ariaLabel: 'Serial No',
                key: 'Id',
                isPadded: true,
                isResizable: true,
                isCollapsible: true,
                isRowHeader: true,
                minWidth: 50,
                maxWidth: 100,
                onRender: (item, index) => (
                    <span>{(this.state.itemsPerPage * this.state.currentPage + index + 1) - (this.state.itemsPerPage)}</span>
                )
            },
            {
                key: 'Title',
                name: 'Title',
                ariaLabel: 'Title',
                isSorted: false,
                sortAscendingAriaLabel: 'Sort A to Z',
                sortDescendingAriaLabel: 'Sort Z to A',
                onColumnClick: this._onColumnClickSort.bind(this),
                isRowHeader: true,
                data: 'string',
                isPadded: true,
                fieldName: 'Title',
                minWidth: 150,
                //maxWidth: 350,
                isResizable: true,
                isCollapsible: false,
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
                key: 'RequestType',
                name: 'Request Type',
                minWidth: 100,
                maxWidth: 150,
                isResizable: true,
                isCollapsible: true,
                onRender: (item) => (
                    <span style={{ color: "#323130", fontSize: 14 }}>
                        {this.currentRequestType === 'Tile Request' ? 'Tile' : 'Collection'}
                    </span>
                )
            },
            {
                key: 'ApprovalStatus',
                name: 'Request Status',
                ariaLabel: 'Request Status',
                sortAscendingAriaLabel: 'Sort A to Z',
                sortDescendingAriaLabel: 'Sort Z to A',
                //onColumnClick: this._onColumnClickSort.bind(this),
                isRowHeader: true,
                data: 'string',
                isPadded: true,
                fieldName: 'ApprovalStatus',
                minWidth: 150,
                maxWidth: 300,
                isResizable: true,
                isCollapsible: true,
            },
            {
                key: 'Action',
                isPadded: true,
                isResizable: true,
                isRowHeader: true,
                isCollapsible: false,
                minWidth: 50,
                maxWidth: 70,
                name: 'Action',
                ariaLabel: 'Action',
                onRender: (item) => (
                    <Link title="Open"
                        onClick={() => this.state.currentRequestType === "Tile Request" ?
                            this._openTileRequestPanel(item) : this._openColRequestPanel(item)
                        }
                    >
                        {
                            this.state.currentRequestStatus !== 'Waiting for Approval' ?
                                (<Icon className="icons" iconName="View" />)
                                : (<Icon className="icons" iconName="Edit" />)
                        }
                    </Link>
                )
            }];
        this.state = {
            ddlRequestTypeValue: 'Tile Request',
            currentRequestType: 'Tile Request',
            currentRequestStatus: 'Waiting for Approval',
            ddlApprovalStatusValue: 'Waiting for Approval',
            detailedListColumns: this.detailedListColumns,
            detailedListItems: [],
            lstApplicationMaster: [],
            lstSearchTilesResult: [],
            isRecordAvailable: true,
            isFilterVisible: true,
            isSortedDesc: false,
            lstColorMaster: [],
            bgColor: "",
            foreColor: "",
            viewScreenControl: false,
            isItemLoaded: false,
            resourceListItems: this.props.resourceListItems,
            currentPage: 1,
            totalPage: 1,
            itemsPerPage: 10,
            detailedListPageItems: [],
            isDataLoaded: false,
            isCorporateCollection: false,

            tNumber: "",
            tNumberId: 0,
            saveControlsErrorMessage: "Please enter mandatory field values.",
            valueRequiredErrorMessage: "This field is mandatory.",
            tileRequestIsPanelOpen: false,
            tileRequestId: 0,
            tileRequestExistingItemId: 0,
            tileRequestTitle: "",
            tileRequestDescription: "",
            tileRequestKeywords: "",
            tileRequestUrlLink: "",
            tileRequestIsUrlValid: true,
            tileRequestOwnerEmail: "",
            tileRequestRequestedAction: "",
            tileRequestRequestedDate: "",
            tileRequestColorCode: 0,
            tileRequestDecisionBy: "",
            tileRequestDecisionDate: "",
            tileRequestDecisionComments: '',
            tileRequestShowMessageBar: false,
            tileRequestAvailableExternal: 1,

            colRequestIsPanelOpen: false,
            colRequestShowMessageBar: false,
            colID: 0,
            colTitle: '',
            //colDefaultMyCollection: 0,
            colDescription: '',
            colPublicCollection: 0,
            colCorporateCollection: 0,
            colStandardOrder: 0,
            colCollectionOwnerId: 0,
            colCollectionOwnerName: '',
            colUnDeletable: 0,
            colRequestedDate: '',
            colApprovalStatus: '',
            colDecisionDate: '',
            colDecisionBy: '',
            colDecisionComments: '',
            colExistingItemID: 0,
            colRequestedAction: '',
            colSelectedTiles: [],

            colMasterCorporateCollection: 0,
            disableColButtons: false,
        };

        //Get Requested data after component mount
        this.getRequestData();
    }

    public componentWillReceiveProps(newProps: IReviewRequestProps) {
        this.setState({
            resourceListItems: newProps.resourceListItems,
        });
    }

    public async componentWillMount(): Promise<void> {
        //get user TNumber
        this.currentUserTNumber = await this.pnpHelper.userProps("TNumber");
        //Get Id from lookup column
        await sp.web.lists.getByTitle(this.lstUserMaster).items.select("Id", "Email").filter(`Title eq '${this.currentUserTNumber}'`).get().then(r => {
            this.currentUserEmail = (r[0].Id !== null || undefined) ? r[0].Email : '';
        });

        //refresh ApplicationMaster an TileRequest data
        this._refreshValidationData();

        let colorCategory: any[] = this.state.lstColorMaster;

        //Color categories exist in state variable
        if (colorCategory.length > 0) {
            colorCategory.map((v: IColorMasterList): void => {
                this.colorOptions.push({ key: v.Id, text: v.Title, selected: false });

            });
            //console.log(colorCategory[0].BgColor + ', ' + colorCategory[0].ForeColor);
            this.setState({ bgColor: colorCategory[0].BgColor, foreColor: colorCategory[0].ForeColor });
        } else {
            //Color categories does not exist in state variable then query SPO list
            try {
                sp.web.lists.getByTitle(this.lstColorMaster).items.orderBy('Id', true).getAll().then((r: any[]): void => {
                    if (r.length > 0) {
                        r.map((v: IColorMasterList) => {
                            this.colorOptions.push({ key: v.Id, text: v.Title });
                        });
                        this.setState({ lstColorMaster: r, bgColor: r[0].BgColor, foreColor: r[0].ForeColor });
                    }
                });
            } catch (error) {
                console.log(error);

                this.errorLogging.logError(this.errTitle, this.errModule, "componentWillMount", error, "componentWillMount");
            }
        }
        //console.log(this.colorOptions);
    }

    public render(): React.ReactElement<IReviewRequestProps> {
        return (
            <React.Fragment >
                <div className={`${styles.reviewRequest}`}>
                    <div className="ms-Grid">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm10 ms-md12 ms-lg12 ms-xl12">
                                <Sticky stickyPosition={StickyPositionType.Header}>
                                    <h2 className="ms-font-m ms-fontWeight-semibold">{this.state.resourceListItems['review_header_text']}</h2>
                                </Sticky>
                            </div>
                        </div>
                        {this.state.isFilterVisible && (
                            <React.Fragment>
                                <hr></hr>
                                <Sticky stickyPosition={StickyPositionType.Header}>
                                    <div className={`ms-Grid-row ${styles.gridRowFilter}`}>
                                        <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg4 ms-xl2">
                                            <label>{this.state.resourceListItems['review_request_type_label']}</label>
                                        </div>
                                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg5 ms-xl3">
                                            <Dropdown
                                                options={[
                                                    { key: 'Tile Request', text: 'Tile Request' },
                                                    { key: 'Collection Request', text: 'Collection Request' },
                                                ]}
                                                // placeholder={'Select option'}
                                                defaultSelectedKey={this.state.ddlRequestTypeValue}
                                                //required={true}
                                                //label={"Request Type:"}
                                                ariaLabel={this.state.resourceListItems['review_request_type_label']}
                                                data-is-focusable={true}
                                                onChange={this.onRequestTypeChange.bind(this)}
                                            ></Dropdown>
                                        </div>
                                        <div className="ms-Grid-col ms-sm5 ms-md5 ms-lg4 ms-xl2">
                                            <label>{this.state.resourceListItems['review_request_status_label']}</label>
                                        </div>
                                        <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg5 ms-xl3">
                                            <Dropdown
                                                options={[
                                                    { key: 'Waiting for Approval', text: 'Waiting for Approval' },
                                                    { key: 'Approved', text: 'Approved' },
                                                    { key: 'Rejected', text: 'Rejected' },
                                                ]}
                                                // placeholder={'Select option'}
                                                defaultSelectedKey={this.state.ddlApprovalStatusValue}
                                                data-is-focusable={true}
                                                //required={true}
                                                //label={"Approval Status:"}
                                                ariaLabel={this.state.resourceListItems['review_request_status_label']}
                                                onChange={this.onApprovalStatusChange.bind(this)}
                                            ></Dropdown>
                                        </div>
                                        <div className="ms-Grid-col">
                                            <PrimaryButton
                                                text={this.state.resourceListItems['review_submit_text']}
                                                onClick={this._onSubmitButtonClicked.bind(this)}
                                                minLength={250}
                                            />
                                        </div>
                                    </div>
                                </Sticky>
                            </React.Fragment>
                        )}
                        <hr></hr>
                        <Sticky stickyPosition={StickyPositionType.Header}>
                            <div className={`ms-Grid-row ${styles.searchRow}`}>
                                <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg8 ms-xl8 detailsListContoll ${styles.paginationDropdown}`}>Show
                                    <Dropdown
                                        label=""
                                        options={this.ddNoOFItems}
                                        styles={this.ddOptionsStyles}
                                        selectedKey={this.state.itemsPerPage}
                                        onChange={this.ddItemsPerPageChanges.bind(this)}
                                        data-is-focusable={true}
                                    />
                                    records
                                </div>
                                {/* <div className={`ms-Grid-col ms-lg4 ms-xl4 ms-hiddenMdDown ${styles.requestTypeText}`}>
                                    <label>{this.state.resourceListItems['review_request_type_label']}:&nbsp;</label>{this.state.currentRequestType}
                                </div> */}
                                {/* <div className="ms-Grid-col ms-sm3 ms-md6 ms-lg3 ms-xl3">
                                <label>{this.state.resourceListItems['review_record_found_text']}: </label>{this.state.recordsFound}
                            </div> */}
                                {/* <div className="ms-Grid-col ms-lg6 ms-xl6 ms-hiddenMdDown">
                                
                                &nbsp;&nbsp;||&nbsp;&nbsp;
                                
                            </div> */}
                                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg4 ms-xl4 detailsListContoll">
                                    <SearchBox
                                        placeholder={this.state.resourceListItems['review_search_placeholder_text']}
                                        iconProps={{ iconName: 'Search' }}
                                        value={this.searchText}
                                        onChange={this._onSeachTextChange.bind(this)}
                                        onSearch={this._onSeachTextChange.bind(this)}
                                        onEscape={this._onSearchClear.bind(this)}
                                        onClear={this._onSearchClear.bind(this)}
                                    />
                                </div>
                            </div>
                        </Sticky>
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12">
                                <DetailsList
                                    setKey="Title"
                                    columns={this.state.detailedListColumns}
                                    items={this.state.detailedListPageItems}
                                    selectionMode={SelectionMode.none}
                                    ariaLabelForGrid={this.state.resourceListItems['review_header_text']}
                                    constrainMode={ConstrainMode.unconstrained}
                                    enableShimmer={!this.state.isRecordAvailable}
                                    isHeaderVisible={true}
                                    //onRenderDetailsHeader={this.onRenderDetailsHeader}
                                    layoutMode={DetailsListLayoutMode.justified}
                                    data-is-scrollable={true}
                                    className={`${styles.detailedList}`}
                                    listProps={{ renderedWindowsAhead: 0, renderedWindowsBehind: 0 }}
                                ></DetailsList>
                                {!this.state.isDataLoaded && (
                                    <Spinner label="Loading content..." hidden={this.state.isDataLoaded}></Spinner>
                                )}

                                {!this.state.isRecordAvailable ?
                                    (
                                        <React.Fragment>
                                            <style dangerouslySetInnerHTML={{
                                                __html: `.ms-DetailsList-contentWrapper { height: 0px !important; }`
                                            }} />
                                            <Label>No records found.</Label>
                                        </React.Fragment>
                                    ) :
                                    (
                                        this.state.isDataLoaded && (
                                            <div className={`${styles.paginationList}`}>
                                                <Pagination
                                                    currentPage={this.state.currentPage}
                                                    totalPages={this.state.totalPage}
                                                    onChange={(page: any) => this.paginationChanged(page)}
                                                />
                                            </div>
                                        )

                                    )
                                }
                            </div>
                        </div>
                    </div>
                </div>
                {/*Tile Request Panel => Start*/}
                <Panel
                    isFooterAtBottom={true}
                    isOpen={this.state.tileRequestIsPanelOpen}
                    onDismiss={() => this._hideTileRequestPanel()}
                    type={PanelType.smallFixedFar}
                    closeButtonAriaLabel="Close"
                    headerText={this.state.resourceListItems['review_tile_header_label']}
                    onRenderFooterContent={this._onRenderTileRequestFooterContent}
                    isLightDismiss={true}
                    className={`${styles.tileReviewRequestPanel}`}
                    headerClassName={styles.headerText}
                // style={{ fontWeight: "bold" }}
                >{
                        this.state.tileRequestIsPanelOpen &&
                        (
                            !this.state.isItemLoaded ?
                                <Spinner label="Loading content..."></Spinner>
                                :
                                !this.state.viewScreenControl ?
                                    (
                                        <div className={`ms-Grid  ${styles.tileReviewRequestPanelContent}`}>
                                            <Label>{this.state.resourceListItems['review_requested_action_label']}: {this.state.tileRequestRequestedAction}</Label>
                                            <TextField
                                                label={this.state.resourceListItems['add_tile_title_label']}
                                                ariaLabel={this.state.resourceListItems['add_tile_title_label']}
                                                placeholder="Tile name"
                                                value={this.state.tileRequestTitle}
                                                onChange={this._onTileRequestTitleChange.bind(this)}
                                                required={true}
                                                //errorMessage={this.state.valueRequiredErrorMessage}
                                                onGetErrorMessage={this.getTileRequestTitleErrorMessage.bind(this)}
                                                validateOnLoad={false}
                                                validateOnFocusIn={true}
                                                validateOnFocusOut={true}
                                                componentRef={this.ctrTitle}
                                                readOnly={this.state.viewScreenControl}
                                                maxLength={255}
                                            />
                                            <TextField
                                                label={this.state.resourceListItems['add_tile_description_label']}
                                                ariaLabel={this.state.resourceListItems['add_tile_description_label']}
                                                placeholder="Tile description"
                                                value={this.state.tileRequestDescription}
                                                onChange={this._onTileRequestDescriptionChange.bind(this)}
                                                required={true}
                                                multiline
                                                rows={3}
                                                onGetErrorMessage={this.getTileRequestDescriptionErrorMessage.bind(
                                                    this
                                                )}
                                                validateOnLoad={false}
                                                validateOnFocusIn={true}
                                                validateOnFocusOut={true}
                                                componentRef={this.ctrDescription}
                                                readOnly={this.state.viewScreenControl}
                                                maxLength={2000}
                                            />
                                            <TextField
                                                label={this.state.resourceListItems['add_tile_keywords_label']}
                                                ariaLabel={this.state.resourceListItems['add_tile_keywords_label']}
                                                placeholder="Comma separated keywords. e.g. Home, Document"
                                                value={this.state.tileRequestKeywords}
                                                onChange={this._onTileRequestKeywordsChange.bind(this)}
                                                required={true}
                                                multiline
                                                rows={3}
                                                onGetErrorMessage={this.getTileRequestKeywordsErrorMessage.bind(
                                                    this
                                                )}
                                                validateOnLoad={false}
                                                validateOnFocusIn={true}
                                                validateOnFocusOut={true}
                                                componentRef={this.ctrKeywords}
                                                readOnly={this.state.viewScreenControl}
                                                maxLength={2000}
                                            />
                                            <TextField
                                                label={this.state.resourceListItems['add_tile_url_link_label']}
                                                ariaLabel={this.state.resourceListItems['add_tile_url_link_label']}
                                                placeholder="e.g. https://pgone.pg.com"
                                                value={this.state.tileRequestUrlLink}
                                                onChange={this._onTileRequestURLLinkChange.bind(this)}
                                                required={true}
                                                multiline
                                                rows={3}
                                                onGetErrorMessage={this.getTileRequestURLLinkErrorMessage.bind(
                                                    this
                                                )}
                                                validateOnLoad={false}
                                                validateOnFocusIn={true}
                                                validateOnFocusOut={true}
                                                componentRef={this.ctrURL}
                                                readOnly={this.state.viewScreenControl}
                                                maxLength={2000}
                                            />
                                            <Link
                                                className={`${styles.linkText}`}
                                                href={this.state.tileRequestUrlLink}
                                                target="_blank"
                                                data-interception="off"
                                                disabled={!this.state.tileRequestIsUrlValid}
                                            >{this.state.resourceListItems['add_tile_try_link_label']}</Link>
                                            <TextField
                                                label={this.state.resourceListItems['add_tile_owneremail_label']}
                                                ariaLabel={this.state.resourceListItems['add_tile_owneremail_label']}
                                                placeholder="e.g. email@pg.com"
                                                value={this.state.tileRequestOwnerEmail}
                                                onChange={this._onTileRequestOwnerEmailChange.bind(this)}
                                                onGetErrorMessage={this.getTileRequestOwnerEmailErrorMessage.bind(this)}
                                                validateOnLoad={false}
                                                validateOnFocusIn={true}
                                                validateOnFocusOut={true}
                                                componentRef={this.ctrOwnerEmail}
                                                readOnly={this.state.viewScreenControl}
                                            />
                                            <div className="ms-Grid-row">
                                                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 ms-xl12">
                                                    {this.state.viewScreenControl ?
                                                        (
                                                            <React.Fragment>
                                                                <Label>{this.state.resourceListItems['add_tile_category_label']}</Label>
                                                                <div style={{
                                                                    backgroundColor: this.state.lstColorMaster[this.selectedColorId - 1].BgColor,
                                                                    color: this.state.lstColorMaster[this.selectedColorId - 1].ForeColor,
                                                                    padding: 6,
                                                                    marginRight: 5,
                                                                    textAlign: "center",
                                                                    border: "1px solid black",
                                                                }}
                                                                >{this.colorOptions.length > 0 && this.colorOptions[this.selectedColorId - 1].text}</div>
                                                            </React.Fragment>
                                                        ) :
                                                        (
                                                            <React.Fragment>
                                                                <style dangerouslySetInnerHTML={{
                                                                    __html: `.ms-Dropdown-container > .ms-Label{
                                                        margin-left: -35px !important;
                                                        }`
                                                                }} />
                                                                <div style={{
                                                                    backgroundColor: this.state.bgColor,
                                                                    color: this.state.foreColor,
                                                                    padding: 6,
                                                                    marginTop: 30,
                                                                    marginRight: 5,
                                                                    width: 15,
                                                                    textAlign: "center",
                                                                    border: "1px solid black",
                                                                    float: "left"
                                                                }}>Aa</div>
                                                                <Dropdown
                                                                    options={this.colorOptions}
                                                                    //placeholder={'Select color option'}
                                                                    selectedKey={this.selectedColorId}
                                                                    label={this.state.resourceListItems['add_tile_category_label']}
                                                                    ariaLabel={this.state.resourceListItems['add_tile_category_label']}
                                                                    required={true}
                                                                    onChange={this._onColorCodeChange.bind(this)}
                                                                    disabled={this.state.viewScreenControl}
                                                                    data-is-focusable={true}
                                                                    errorMessage={this.selectedColorId === 0 && this.showError && this.state.resourceListItems['required_field_validation_message']}
                                                                ></Dropdown>
                                                            </React.Fragment>
                                                        )
                                                    }
                                                </div>
                                            </div>
                                            <Toggle
                                                label={this.state.resourceListItems['add_tile_available_external_label']}
                                                onText="Yes"
                                                offText="No"
                                                checked={this.state.tileRequestAvailableExternal == 1 ? false : true}
                                                onChange={this._onTileRequestAvailableExternalChange.bind(this)}
                                            />
                                            <Label>{this.state.resourceListItems['review_requested_date_label']}</Label>
                                            <div>{this.state.tileRequestRequestedDate}</div>
                                            {this.state.viewScreenControl && (
                                                <React.Fragment>
                                                    <Label>{this.state.resourceListItems['review_decision_date_label']}</Label>
                                                    <div>{this.state.tileRequestDecisionDate}</div>
                                                    <Label>{this.state.resourceListItems['review_decision_by_label']}</Label>
                                                    <div>{this.state.tileRequestDecisionBy}</div>
                                                </React.Fragment>
                                            )}
                                            <TextField
                                                label={this.state.resourceListItems['review_decision_comments_label']}
                                                ariaLabel={this.state.resourceListItems['review_decision_comments_label']}
                                                placeholder="Decision Comments"
                                                value={this.state.tileRequestDecisionComments}
                                                onChange={this._onTileRequestDecisionCommentsChange.bind(this)}
                                                required={true}
                                                multiline
                                                rows={3}
                                                onGetErrorMessage={this.getTileRequestDecisionCommentsErrorMessage.bind(
                                                    this
                                                )}
                                                validateOnLoad={false}
                                                validateOnFocusIn={true}
                                                validateOnFocusOut={true}
                                                componentRef={this.ctrDecisionComments}
                                                readOnly={this.state.viewScreenControl}
                                                maxLength={2000}
                                            />
                                        </div>
                                    ) :
                                    (
                                        <div className={`ms-Grid  ${styles.tileReviewRequestPanelContent}`}>
                                            <Label>{this.state.resourceListItems['review_requested_action_label']}</Label>
                                            <div>{this.state.tileRequestRequestedAction}</div>
                                            <Label>{this.state.resourceListItems['add_tile_title_label']}</Label>
                                            <div>{this.state.tileRequestTitle}</div>
                                            <Label>{this.state.resourceListItems['add_tile_description_label']}</Label>
                                            <div>{this.state.tileRequestDescription}</div>
                                            <Label>{this.state.resourceListItems['add_tile_keywords_label']}</Label>
                                            <div>{this.state.tileRequestKeywords}</div>
                                            <Label>{this.state.resourceListItems['add_tile_url_link_label']}</Label>
                                            <div>{this.state.tileRequestUrlLink}</div>
                                            <Link
                                                className={`${styles.linkText}`}
                                                href={this.state.tileRequestUrlLink}
                                                target="_blank"
                                                disabled={!this.state.tileRequestIsUrlValid}
                                            >{this.state.resourceListItems['add_tile_try_link_label']}</Link>
                                            <Label>{this.state.resourceListItems['add_tile_owneremail_label']}</Label>
                                            <div>{this.state.tileRequestOwnerEmail}</div>
                                            <Label>{this.state.resourceListItems['add_tile_category_label']}</Label>
                                            <div style={{
                                                backgroundColor: this.state.lstColorMaster[this.selectedColorId - 1].BgColor,
                                                color: this.state.lstColorMaster[this.selectedColorId - 1].ForeColor,
                                                padding: 6,
                                                marginRight: 5,
                                                textAlign: "center",
                                                border: "1px solid black",
                                            }}
                                            >{this.colorOptions.length > 0 && this.colorOptions[this.selectedColorId].text}</div>
                                            <Label>{this.state.resourceListItems['add_tile_available_external_label']}</Label>
                                            <div>{this.state.tileRequestAvailableExternal === 1 ? 'No' : 'Yes'}</div>
                                            <Label>{this.state.resourceListItems['review_requested_date_label']}</Label>
                                            <div>{this.state.tileRequestRequestedDate}</div>
                                            <Label>{this.state.resourceListItems['review_decision_date_label']}</Label>
                                            <div>{this.state.tileRequestDecisionDate}</div>
                                            <Label>{this.state.resourceListItems['review_decision_by_label']}</Label>
                                            <div>{this.state.tileRequestDecisionBy}</div>
                                            <Label>{this.state.resourceListItems['review_decision_comments_label']}</Label>
                                            <div>{this.state.tileRequestDecisionComments}</div>
                                        </div>
                                    )
                        )
                    }
                </Panel>
                {/*Tile Request Panel => End*/}
                {/*Collection Request Panel => Start*/}
                <Panel
                    isFooterAtBottom={true}
                    isOpen={this.state.colRequestIsPanelOpen}
                    onDismiss={() => this._hideColRequestPanel()}
                    type={PanelType.smallFixedFar}
                    closeButtonAriaLabel="Close"
                    headerText={this.state.resourceListItems['review_collection_header_label']}
                    onRenderFooterContent={this._onRenderColRequestFooterContent}
                    isLightDismiss={true}
                    className={`${styles.createCollectionPanel}`}
                    headerClassName={styles.headerText}
                >
                    {
                        this.state.colRequestIsPanelOpen &&
                        (
                            !this.state.isItemLoaded ?
                                <Spinner label="Loading content..."></Spinner>
                                :
                                !this.state.viewScreenControl ?
                                    (
                                        <div className={`ms-Grid  ${styles.createCollectionPanelContent}`}>
                                            <div className={`ms-Grid-row`}>
                                                <div className={`ms-Grid-row`}>
                                                    <div className={`ms-Grid-col sm-12 md-12 lg-6 xl-6`}>
                                                        <Label>{this.state.resourceListItems['review_requested_action_label']}</Label>
                                                        <div>{this.state.colRequestedAction}</div>
                                                    </div>
                                                    <div className={`ms-Grid-col sm-12 md-12 lg-6 xl-6`}>
                                                        <Label>{this.state.resourceListItems['review_collection_type_label']}</Label>
                                                        <div>{this.state.colPublicCollection === 1 ? "Public" : "Public to Private"}</div>
                                                    </div>
                                                </div>
                                                <TextField
                                                    label={this.state.resourceListItems['add_tile_title_label']}
                                                    ariaLabel={this.state.resourceListItems['add_tile_title_label']}
                                                    placeholder="Collection name"
                                                    value={this.state.colTitle}
                                                    onChange={this._onColRequestTitleChange.bind(this)}
                                                    required={true}
                                                    //errorMessage={this.state.valueRequiredErrorMessage}
                                                    onGetErrorMessage={this.getColRequestTitleErrorMessage.bind(this)}
                                                    validateOnLoad={false}
                                                    validateOnFocusIn={true}
                                                    validateOnFocusOut={true}
                                                    componentRef={this.ctrColTitle}
                                                    readOnly={this.state.viewScreenControl}
                                                    maxLength={255}
                                                />
                                                <TextField
                                                    label={this.state.resourceListItems['add_tile_description_label']}
                                                    ariaLabel={this.state.resourceListItems['add_tile_description_label']}
                                                    placeholder="Collection description"
                                                    value={this.state.colDescription}
                                                    onChange={this._onColRequestDescriptionChange.bind(this)}
                                                    required={true}
                                                    multiline
                                                    rows={3}
                                                    onGetErrorMessage={this.getColRequestDescriptionErrorMessage.bind(this)}
                                                    validateOnLoad={false}
                                                    validateOnFocusIn={true}
                                                    validateOnFocusOut={true}
                                                    componentRef={this.ctrColDescription}
                                                    readOnly={this.state.viewScreenControl}
                                                    maxLength={2000}
                                                />
                                                <Label required>{this.state.resourceListItems['review_collection_selected_tile_label']}</Label>
                                                <div title="selectedTiles" style={{ maxHeight: 150, overflow: "auto" }}>
                                                    {
                                                        this.state.lstApplicationMaster.map((applicationTile: any) =>
                                                            this.selectedTiles.some((selectedTile: any) => {
                                                                return selectedTile.ApplicationIDId == applicationTile.Id;
                                                            }) && (
                                                                <div style={{
                                                                    padding: 5,
                                                                    margin: 3,
                                                                    minWidth: 30,
                                                                    float: "left",
                                                                    borderRadius: 5,
                                                                    textAlign: "center",
                                                                    backgroundColor: this.state.lstColorMaster[applicationTile.ColorCodeId - 1].BgColor,
                                                                    color: this.state.lstColorMaster[applicationTile.ColorCodeId - 1].ForeColor,
                                                                }}>
                                                                    {applicationTile.Title}
                                                                </div>
                                                            )
                                                        )
                                                    }
                                                </div>
                                                <div className={`ms-Grid-row`}>
                                                    <div className={`ms-Grid-col sm-6 md-6 lg-6 xl-6`}>
                                                        <Toggle
                                                            label={this.state.resourceListItems['review_collection_corporate_label']}
                                                            onText="Yes"
                                                            offText="No"
                                                            checked={this.state.colCorporateCollection == 1 ? true : false}
                                                            onChange={this._onColRequestCorporateCollectionChange.bind(this)}
                                                            disabled={this.state.isCorporateCollection}
                                                        />
                                                    </div>
                                                    <div className={`ms-Grid-col sm-6 md-6 lg-6 xl-6`}>
                                                        <Toggle
                                                            label={this.state.resourceListItems['review_collection_undeletable_label']}
                                                            onText="Yes"
                                                            offText="No"
                                                            checked={this.state.colUnDeletable == 1 ? true : false}
                                                            onChange={this._onColRequestUndeletableChange.bind(this)}
                                                            disabled={this.state.viewScreenControl}
                                                        />
                                                    </div>
                                                </div>
                                                <Label>{this.state.resourceListItems['review_requested_date_label']}</Label>
                                                <div>{this.state.colRequestedDate}</div>
                                                {this.state.viewScreenControl && (
                                                    <React.Fragment>
                                                        <Label>{this.state.resourceListItems['review_decision_date_label']}</Label>
                                                        <div>{this.state.colDecisionDate}</div>
                                                    </React.Fragment>
                                                )}
                                                <Label>{this.state.resourceListItems['review_collection_owner_label']}</Label>
                                                <div>{this.state.colCollectionOwnerName}</div>
                                                <TextField
                                                    label={this.state.resourceListItems['review_decision_comments_label']}
                                                    ariaLabel={this.state.resourceListItems['review_decision_comments_label']}
                                                    placeholder="Decision Comments"
                                                    value={this.state.colDecisionComments}
                                                    onChange={this._onColRequestDecisionCommentsChange.bind(this)}
                                                    required={true}
                                                    multiline
                                                    rows={3}
                                                    onGetErrorMessage={this.getColRequestDecisionCommentsErrorMessage.bind(this)}
                                                    validateOnLoad={false}
                                                    validateOnFocusIn={true}
                                                    validateOnFocusOut={true}
                                                    componentRef={this.ctrColDecisionComments}
                                                    readOnly={this.state.viewScreenControl}
                                                    maxLength={2000}
                                                />
                                            </div>

                                        </div>
                                    ) :
                                    (
                                        <div className={`ms-Grid  ${styles.createCollectionPanelContent}`}>
                                            <div className={`ms-Grid-row`}>
                                                <Label>{this.state.resourceListItems['review_requested_action_label']}</Label>
                                                <div>{this.state.colRequestedAction}</div>
                                                <Label>{this.state.resourceListItems['review_collection_type_label']}</Label>
                                                <div>{this.state.colPublicCollection == 1 && "Public"}</div>
                                                <Label>{this.state.resourceListItems['add_tile_title_label']}</Label>
                                                <div>{this.state.colTitle}</div>
                                                <Label>{this.state.resourceListItems['add_tile_description_label']}</Label>
                                                <div>{this.state.colDescription}</div>

                                                <Label required>{this.state.resourceListItems['review_collection_selected_tile_label']}</Label>
                                                <div title="selectedTiles" style={{ maxHeight: 150, overflow: "auto" }}>
                                                    {
                                                        this.state.lstApplicationMaster.map((applicationTile: any) =>
                                                            this.selectedTiles.some((selectedTile: any) => {
                                                                return selectedTile.ApplicationIDId == applicationTile.Id;
                                                            }) && (
                                                                <div style={{
                                                                    padding: 5,
                                                                    margin: 3,
                                                                    minWidth: 30,
                                                                    float: "left",
                                                                    borderRadius: 5,
                                                                    textAlign: "center",
                                                                    backgroundColor: this.state.lstColorMaster[applicationTile.ColorCodeId - 1].BgColor,
                                                                    color: this.state.lstColorMaster[applicationTile.ColorCodeId - 1].ForeColor,
                                                                }}>
                                                                    {applicationTile.Title}
                                                                </div>
                                                            )
                                                        )
                                                    }
                                                </div>

                                                <Label>{this.state.resourceListItems['review_collection_corporate_label']}</Label>
                                                <div>{this.state.colCorporateCollection == 1 ? 'Yes' : 'No'}</div>
                                                <Label>{this.state.resourceListItems['review_collection_undeletable_label']}</Label>
                                                <div>{this.state.colUnDeletable == 1 ? 'Yes' : 'No'}</div>
                                                <Label>{this.state.resourceListItems['review_requested_date_label']}</Label>
                                                <div>{this.state.colRequestedDate}</div>
                                                <Label>{this.state.resourceListItems['review_collection_owner_label']}</Label>
                                                <div>{this.state.colCollectionOwnerName}</div>
                                                <Label>{this.state.resourceListItems['review_decision_date_label']}</Label>
                                                <div>{this.state.colDecisionDate}</div>
                                                <Label>{this.state.resourceListItems['review_decision_by_label']}</Label>
                                                <div>{this.state.colDecisionBy}</div>
                                                <Label>{this.state.resourceListItems['review_decision_comments_label']}</Label>
                                                <div>{this.state.colDecisionComments}</div>
                                            </div>

                                        </div>
                                    )
                        )
                    }
                </Panel>
                {/*Collection Request Panel => End*/}
            </React.Fragment >
        );
    }

    //#region Review Request

    //#region Click Events
    //On Submit button Click
    private _onSubmitButtonClicked = () => {
        try {
            //Get Requested data on submit and clear existing items from items array
            this.setState({
                currentRequestType: this.state.ddlRequestTypeValue,
                currentRequestStatus: this.state.ddlApprovalStatusValue,
                detailedListItems: [],
                isRecordAvailable: false,
                recordsFound: 0,
                isDataLoaded: false,
            });
            this.currentRequestType = this.state.ddlRequestTypeValue;
            this.currentRequestStatus = this.state.ddlApprovalStatusValue;
            this.searchData = [];
            this.getRequestData();
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_onSubmitButtonClicked");
        }

    }

    //call this function when tile request approved or rejected
    //this function will refresh the datasets for validation
    private _refreshValidationData = () => {
        try {
            //populate application master data
            sp.web.lists.getByTitle(this.lstTileMaster).items.getAll().then(r => {
                //console.log(r);
                this.setState({ lstApplicationMaster: r, lstSearchTilesResult: r });
            });
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_onSubmitButtonClicked");
        }

    }

    //Get Request data based on request type
    private getRequestData = (): Promise<any> => {
        return new Promise<any>(
            async (
                resolve: (items: any[]) => void,
                reject: (error: any) => void
            ): Promise<void> => {
                try {

                    //console.log("isDataLoaded: false =>" + this.state.isDataLoaded);
                    let requestType = this.currentRequestType;
                    let listName = requestType === "Tile Request" ? this.lstTileRequest : this.lstCollectionRequest;
                    let itemCount = (await sp.web.lists.getByTitle(listName).items.getAll()).length;
                    //console.log("itemcount:" + itemCount);
                    let spCall: Promise<any[]> = itemCount > 5000 ? (
                        sp.web.lists
                            .getByTitle(listName)
                            .items
                            .getAll()
                            .then((r): any[] => {
                                let results: any[] = [];
                                if (r.length > 0) {
                                    results = r
                                        .filter(v => v.ApprovalStatus === this.currentRequestStatus)
                                        .sort((a, b) => { return b.ID - a.ID; });
                                }
                                return results;
                            }, (error: any): void => {
                                reject(error);
                            })
                    ) : (
                            sp.web.lists
                                .getByTitle(listName)
                                .items
                                .filter(`ApprovalStatus eq '${this.currentRequestStatus}'`)
                                .orderBy("Id", false)
                                .get()
                                .then((r): any[] => {
                                    return r;
                                }, (error: any): void => {
                                    reject(error);
                                })
                        );

                    //Set state with variable
                    spCall.then((results) => {
                        //set detailed list items in state
                        this.setState({
                            detailedListItems: results,
                            isRecordAvailable: true,
                            recordsFound: results.length,
                        });
                        this.searchData = results;
                        //Pagination setup
                        this.paginationConfig(results, 1);

                        //remove 
                        this.setState({ isDataLoaded: true });
                        //console.log("isDataLoaded: true =>" + this.state.isDataLoaded);
                    });


                } catch (error) {

                    reject(error);
                    this.errorLogging.logError(this.errTitle, this.errModule, "", error, "getRequestData");
                }
            }
        );
    }

    //Dropdown Items per page changed
    private ddItemsPerPageChanges =
        (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
            let varkey = parseInt(item.key.toString());
            //console.log(varkey);
            this.setState({
                itemsPerPage: varkey,
            });
            //call pagination
            this.paginationConfig(this.state.detailedListItems, 1, varkey);
        }

    private paginationChanged = (page: any): void => {
        //call pagination
        this.paginationConfig(this.state.detailedListItems, page);
    }

    //Search text based filter
    private _onSeachTextChange = (searchValue: string) => {
        try {
            let resultData: any[] = [];
            this.searchText = searchValue;
            this.searchData.map((item: any): void => {
                if (
                    String(item["Title"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(item["RequestedAction"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(item["SearchKeywords"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(item["Description"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1
                ) {
                    resultData.push(item);
                }
            });
            //set result set into state
            this.setState({ detailedListItems: resultData });
            //call pagination
            this.paginationConfig(resultData);
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_onSeachTextChange");
        }
    }

    //Search text clear
    private _onSearchClear = () => {
        this.searchText = "";
        this.setState({ detailedListItems: this.searchData });
    }

    //Column sort function
    private _onColumnClickSort = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        try {
            const { detailedListColumns } = this.state;
            let { detailedListItems } = this.state;

            let isSortedDescending = column.isSortedDescending;
            // If we've sorted this column, flip it.
            if (column.isSorted) {
                isSortedDescending = !isSortedDescending;
            }
            // Sort the items.
            detailedListItems = this._copyAndSort(detailedListItems, column.fieldName!, isSortedDescending);

            // Reset the items and columns to match the state.
            this.setState({
                detailedListItems: detailedListItems,
                detailedListColumns: detailedListColumns.map(col => {
                    col.isSorted = col.key === column.key;
                    if (col.isSorted) {
                        col.isSortedDescending = isSortedDescending;
                    }
                    return col;
                }),
            });

            //call pagination
            this.paginationConfig(detailedListItems, 1);
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_onColumnClickSort");
        }
    }

    private _copyAndSort<T>(items: any[], columnKey: string, isSortedDescending?: boolean): any[] {
        const key = columnKey as keyof any;
        //return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
        return items.slice(0).sort((a: any, b: any) => ((isSortedDescending ? a[key].toString().toUpperCase() < b[key].toString().toUpperCase() : a[key].toString().toUpperCase() > b[key].toString().toUpperCase()) ? 1 : -1));

    }

    //Detailed list Pagination logic
    private paginationConfig = (data: any[], currentPage?: number, itemsPerPage?: any) => {
        try {
            let vCurrentPage = currentPage === undefined || null ? 1 : currentPage;
            let vItemsPerPage = itemsPerPage === undefined || null ? this.state.itemsPerPage : itemsPerPage;
            if (data.length !== 0) {
                this.setState({
                    currentPage: vCurrentPage,
                    totalPage: Math.ceil(data.length / vItemsPerPage),
                    detailedListPageItems: data.slice(vItemsPerPage * vCurrentPage - vItemsPerPage, vItemsPerPage * vCurrentPage),
                    isRecordAvailable: true,
                });
            } else {
                this.setState({
                    currentPage: 1,
                    totalPage: 1,
                    detailedListPageItems: [],
                    isRecordAvailable: false,
                });
            }
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "paginationConfig");
        }
    }
    //#endregion

    //#region Capture Values
    //Set Request Type state value 
    private onRequestTypeChange =
        (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
            this.setState({ ddlRequestTypeValue: item.text });
        }
    //Set Request Status state value
    private onApprovalStatusChange =
        (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
            this.setState({ ddlApprovalStatusValue: item.text });
        }

    //#endregion

    //#endregion

    //#region tile request Events and Validation

    //#region Tile Request Other events
    //Called when user clicked on "Add a tile" option from left pane
    private _openTileRequestPanel = async (item: ITileRequest) => {
        try {
            //refresh application request data and ApplicationMasterDate
            this._refreshValidationData();
            this.setState({ isItemLoaded: false });
            //let colorCodeItemID = this.colorOptions.length > 0 && (parseInt(this.colorOptions[0].key.toString()));
            this.selectedColorId = isNaN(item.ColorCodeId) || (item.ColorCodeId === null || undefined) ? 1 : item.ColorCodeId;
            //console.log(this.selectedColorId);

            let colorMaster = this.state.lstColorMaster;
            this.setState({
                tileRequestId: item.Id,
                tileRequestExistingItemId: item.ExistingItemID,
                tileRequestTitle: item.Title,
                tileRequestDescription: item.Description,
                tileRequestKeywords: item.SearchKeywords,
                tileRequestUrlLink: item.Link,
                tileRequestOwnerEmail: item.OwnerEmail,
                tileRequestColorCode: this.selectedColorId,
                bgColor: colorMaster.length > 0 && colorMaster[this.selectedColorId - 1].BgColor,
                foreColor: colorMaster.length > 0 && colorMaster[this.selectedColorId - 1].ForeColor,
                tileRequestRequestedDate: item.RequestedDate != null || undefined
                    ? new Date(item.RequestedDate.toString()).toLocaleString() : "",
                tileRequestRequestedAction: item.RequestedAction,
                tileRequestDecisionBy: item.DecisionBy,
                tileRequestDecisionDate: item.DecisionDate != null || undefined
                    ? new Date(item.DecisionDate.toString()).toLocaleString() : "",
                tileRequestDecisionComments: stringIsNullOrEmpty(item.DecisionComments) ? "" : item.DecisionComments,
                tileRequestIsUrlValid: true,
                viewScreenControl: item.ApprovalStatus != 'Waiting for Approval' ? true : false,
                tileRequestIsPanelOpen: true,
                isItemLoaded: true,
                tileRequestAvailableExternal: item.AvailableExternal === null ? 1 : item.AvailableExternal,
            });
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_openTileRequestPanel");
        }

    }

    //to close panel
    private _hideTileRequestPanel = async (): Promise<void> => {
        try {
            if (this.isDataRefresh) {
                //refresh detailed data
                this.getRequestData().then((): void => {
                    //clear controls state
                    this.clearTileReuqestControlsState();
                });

            } else {
                this.clearTileReuqestControlsState();
            }

            this.setState({ tileRequestIsPanelOpen: false });
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_hideTileRequestPanel");
        }
    }

    //Populate panel footer content with "Save" and "Cancel" buttons
    private _onRenderTileRequestFooterContent = (): JSX.Element => {
        return (
            <React.Fragment>

                {this.state.tileRequestShowMessageBar && (
                    <div>
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={false}
                            onDismiss={this._closeTileRequestMessageBar.bind(this)}
                            dismissButtonAriaLabel="Close"
                        >
                            {this.state.saveControlsErrorMessage}
                        </MessageBar>
                        <br />
                    </div>
                )}
                {!this.state.viewScreenControl && (
                    <div>
                        <DefaultButton
                            className={`${styles.approveBtn}`}
                            text={this.state.resourceListItems['btn_approve_text']}
                            iconProps={{ iconName: 'Checkmark' }}
                            onClick={this._saveTileRequest.bind(this, "Approved")}

                        />
                        &nbsp;
                        <DefaultButton
                            className={`${styles.rejectBtn}`}
                            text={this.state.resourceListItems['btn_reject_text']}
                            iconProps={{ iconName: 'Clear' }}
                            onClick={this._saveTileRequest.bind(this, "Rejected")}
                        />
                        <br />
                    </div>

                )}
            </React.Fragment >
        );
    }

    // User clicks on save button, new item will create in "ApplicationRequests" list
    private _saveTileRequest = async (status: string) => {
        try {
            //console.log(status + ", " + this.currentUserEmail);
            let isFormValid = this.validateTileReuqestControls(status);
            //console.log(isFormValid);
            if (isFormValid) {
                //this.setState({ tileRequestShowMessageBar: false });

                let itemTileRequest: ITileRequest;
                if (status === "Approved") {
                    if (this.state.tileRequestRequestedAction !== 'Deletion') {
                        itemTileRequest = {
                            Title: this.state.tileRequestTitle,
                            Description: this.state.tileRequestDescription,
                            SearchKeywords: this.state.tileRequestKeywords,
                            Link: this.state.tileRequestUrlLink,
                            OwnerEmail: this.state.tileRequestOwnerEmail,
                            ColorCodeId: this.state.tileRequestColorCode,
                            AvailableExternal: this.state.tileRequestAvailableExternal,
                            //RequestedById: this.state.tNumberId,
                            DecisionBy: this.currentUserEmail,
                            DecisionComments: this.state.tileRequestDecisionComments,
                            ApprovalStatus: status,
                            DecisionDate: new Date(),
                        };
                    } else {
                        itemTileRequest = {
                            DecisionBy: this.currentUserEmail,
                            DecisionComments: this.state.tileRequestDecisionComments,
                            ApprovalStatus: status,
                            DecisionDate: new Date(),
                        };
                    }
                } else {
                    itemTileRequest = {
                        DecisionBy: this.currentUserEmail,
                        DecisionComments: this.state.tileRequestDecisionComments,
                        ApprovalStatus: status,
                        DecisionDate: new Date(),
                    };
                }

                //console.log(itemTileRequest);
                //Update existing request data to ApplcationRequest List
                try {
                    sp.web.lists.getByTitle(this.lstTileRequest)
                        .items.getById(this.state.tileRequestId).update(itemTileRequest).then(r => {
                            //console.log(r);
                            //Hide panel
                            this.isDataRefresh = true;
                            this._hideTileRequestPanel();
                            this.props.callBackForRequestSection("Request reviewed sucessfully.");
                        });
                } catch (error) {
                    //console.log(error);
                }

            } else {
                this.ctrTitle.current.focus();
                this.ctrDescription.current.focus();
                this.ctrKeywords.current.focus();
                this.ctrURL.current.focus();
                this.ctrOwnerEmail.current.focus();
                this.ctrDecisionComments.current.focus();
                //show error message
                this.showError = this.selectedColorId === 0 ? true : false;
                //this.setState({ tileRequestShowMessageBar: true });
            }
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_saveTileRequest");
        }
    }

    //close MessageBar
    private _closeTileRequestMessageBar = () => {
        this.setState({ tileRequestShowMessageBar: false });
    }

    //validate controls state
    private validateTileReuqestControls = (status: string): boolean => {
        let isValid = true;
        if (status === "Approved") {
            if (this.getTileRequestTitleErrorMessage(this.state.tileRequestTitle) !== "")
                isValid = false;
            if (this.getTileRequestDescriptionErrorMessage(this.state.tileRequestDescription) !== "")
                isValid = false;
            if (this.getTileRequestKeywordsErrorMessage(this.state.tileRequestKeywords) !== "")
                isValid = false;
            if (this.getTileRequestURLLinkErrorMessage(this.state.tileRequestUrlLink) !== "") isValid = false;
            if (!this.state.tileRequestIsUrlValid)
                isValid = false;
            if (this.getTileRequestOwnerEmailErrorMessage(this.state.tileRequestOwnerEmail) !== "")
                isValid = false;
            if (this.state.tileRequestColorCode === 0)
                isValid = false;
            if (this.showError)
                isValid = false;
            if (this.getTileRequestDecisionCommentsErrorMessage(this.state.tileRequestDecisionComments) !== "")
                isValid = false;
        } else {
            if (this.getTileRequestDecisionCommentsErrorMessage(this.state.tileRequestDecisionComments) !== "")
                isValid = false;
        }
        return isValid;
    }

    //clear controls state after panel close
    private clearTileReuqestControlsState = (): void => {
        this.setState({
            tileRequestId: 0,
            tileRequestTitle: "",
            tileRequestDescription: "",
            tileRequestKeywords: "",
            tileRequestUrlLink: "",
            tileRequestOwnerEmail: "",
            tileRequestRequestedDate: "",
            tileRequestColorCode: 0,
            tileRequestAvailableExternal: 1,
            bgColor: '',
            foreColor: '',
            tileRequestShowMessageBar: false,
            tileRequestDecisionBy: '',
            tileRequestDecisionDate: '',
            tileRequestDecisionComments: '',
            tileRequestIsUrlValid: false,
            viewScreenControl: false,
            isItemLoaded: false,
            //tileRequestIsPanelOpen: false,
        });

        this.selectedColorId = 1;
        this.showError = false;
        this.isDataRefresh = false;
    }
    //#endregion

    //#region Validation
    //Validate Title text
    private getTileRequestTitleErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message'] : value.length > 255
                ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //Validate Description text
    private getTileRequestDescriptionErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message'] : value.length > 63999
                ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //Validate Keywords text
    private getTileRequestKeywordsErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message'] : value.length > 63999
                ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //Validate URLLink text
    private getTileRequestURLLinkErrorMessage = (value: string): string => {
        value = value.trim();
        let urlPattern = /(http(s)?:\/\/.)(www\.)?[-a-zA-Z0-9@:%._\+~#=]{2,63984}\.[a-z]{2,6}\b([-a-zA-Z0-9@:%_\+.~#?&\/=]*)/;
        let errMsg =
            stringIsNullOrEmpty(value)
                ? this.state.resourceListItems['required_field_validation_message']
                : value.length > 63999 ? this.state.resourceListItems['validation_maximum_characters_text']
                    : !urlPattern.test(value) ? this.state.resourceListItems['validation_invalid_url_text']
                        : this.state.lstApplicationMaster.some(tile => {
                            if (tile.IsActive) {
                                if (this.state.tileRequestRequestedAction !== 'Addition') {
                                    if (parseInt(tile.ID) !== this.state.tileRequestExistingItemId) {
                                        return tile.Link.toLowerCase() === value.toLowerCase();
                                    }
                                } else {
                                    return tile.Link.toLowerCase() === value.toLowerCase();
                                }
                            }
                        }) ? this.state.resourceListItems['validation_url_exist_text']
                            : this.state.lstApplicationMaster.some(tile => {
                                if (!tile.IsActive) {
                                    if (this.state.tileRequestRequestedAction !== 'Addition') {
                                        if (parseInt(tile.ID) !== this.state.tileRequestExistingItemId) {
                                            return tile.Link.toLowerCase() === value.toLowerCase();
                                        }
                                    } else {
                                        return tile.Link.toLowerCase() === value.toLowerCase();
                                    }
                                }
                            }) ? this.state.resourceListItems['validation_url_exist_disable_text']
                                : (this.state.currentRequestType === 'Tile Request' && this.searchData.some(tile => {
                                    if (tile.ID !== this.state.tileRequestId)
                                        return tile.Link.toLowerCase() === value.toLowerCase();
                                })) ? this.state.resourceListItems['add_tile_request_exist_msg']
                                    : "";

        //to disable Test URL
        this.setState({
            tileRequestIsUrlValid: stringIsNullOrEmpty(errMsg) ? true : false,
        });
        return errMsg;
    }
    //Validate OwnerEmail text
    private getTileRequestOwnerEmailErrorMessage = (value: string): string => {
        value = value.trim();
        let emailPattern = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@(pg.com)$/;
        let errMsg =
            stringIsNullOrEmpty(value)
                ? this.state.resourceListItems['required_field_validation_message']
                : value.length > 500
                    ? this.state.resourceListItems['validation_maximum_characters_text']
                    : emailPattern.test(value) ? ""
                        : this.state.resourceListItems['validation_invalid_email'];
        return errMsg;
    }

    private getTileRequestDecisionCommentsErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message']
            : value.length > 500 ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //#endregion

    //#region Capture text
    //capture Title text
    private _onTileRequestTitleChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ): void => {
        this.setState({ tileRequestTitle: newValue });
    }

    //capture Description text
    private _onTileRequestDescriptionChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ tileRequestDescription: newValue });
    }

    //capture Keywords text
    private _onTileRequestKeywordsChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ tileRequestKeywords: newValue });
    }

    //capture URLLink text
    private _onTileRequestURLLinkChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        // let val: string =
        //   newValue.indexOf("http") == -1 ? "http://" + newValue : newValue;
        this.setState({ tileRequestUrlLink: newValue });
    }

    //capture OwnerEmail text
    private _onTileRequestOwnerEmailChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ tileRequestOwnerEmail: newValue });
    }

    //capture color code
    private _onColorCodeChange = (
        event: React.FormEvent<HTMLDivElement>,
        item: IDropdownOption
    ) => {
        this.selectedColorId = parseInt(item.key.toString());
        // let colorCode = parseInt(item.key.toString());
        //console.log("lstColorMaster:", this.state.lstColorMaster);
        if (this.selectedColorId > 0) {
            this.setState({
                tileRequestColorCode: this.selectedColorId,
                bgColor: this.state.lstColorMaster[this.selectedColorId - 1].BgColor,
                foreColor: this.state.lstColorMaster[this.selectedColorId - 1].ForeColor,
            });
        } else {
            this.setState({
                tileRequestColorCode: this.selectedColorId,
                bgColor: "#FFFFFF",
                foreColor: "#000000",
            });
        }

        //show error message
        this.showError = this.selectedColorId === 0 ? true : false;

    }

    //capture Available External text
    private _onTileRequestAvailableExternalChange = (
        event: React.MouseEvent<HTMLElement>,
        newText: boolean
    ) => {
        this.setState({ tileRequestAvailableExternal: newText ? 0 : 1 });
    }

    //capture DecisionComments text
    private _onTileRequestDecisionCommentsChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ tileRequestDecisionComments: newValue });
    }
    //#endregion

    //#endregion TileRequest Events and Validation Ends

    //#region Collection Request Events and Validations

    //#region other events
    //Called when user clicked on "Add a tile" option from left pane
    private _openColRequestPanel = async (item: ICollectionRequest) => {
        try {
            //open panel
            this.setState({
                colRequestIsPanelOpen: true,
                isItemLoaded: false,
            });

            //get collection owner email id
            await sp.web.lists.getByTitle(this.lstUserMaster)
                .items
                .getById(item.CollectionOwnerId)
                .select("Id", "Email")
                .get()
                .then(r => {
                    //console.log(r);
                    this.setState({ colCollectionOwnerName: (r.Id !== null || undefined) ? r.Email : '' });
                });

            let listName: string = item.RequestedAction === "Addition" ? this.lstColAppMatrixRequests : this.lstColAppMatrix;
            let varFilter: string = item.RequestedAction === "Addition" ? `CollectionRequestIDId eq '${item.ID}'` : `CollectionIDId eq '${item.ExistingItemID}'`;

            //get application tiles for collection request
            await sp.web.lists
                .getByTitle(listName)
                .items.filter(varFilter)
                .select("ApplicationIDId").get().then(r => {
                    this.setState({ colSelectedTiles: r });
                    //push selected tiles to array
                    r.map((value) => {
                        this.selectedTiles.push(value);
                        //console.log("selected tiles:" + value.ApplicationIDId);
                    });

                });

            this.setState({
                //colRequestIsPanelOpen: true,
                colRequestShowMessageBar: false,
                colID: item.ID,
                colTitle: item.Title,
                //colDefaultMyCollection: item.DefaultMyCollection,
                colDescription: item.Description,
                colPublicCollection: item.PublicCollection,
                colCorporateCollection: item.CorporateCollection,
                colStandardOrder: item.StandardOrder,
                colCollectionOwnerId: item.CollectionOwnerId,
                colUnDeletable: item.UnDeletable,
                colRequestedDate: item.RequestedDate != null || undefined
                    ? new Date(item.RequestedDate.toString()).toLocaleString() : "",
                colApprovalStatus: item.ApprovalStatus,
                colDecisionBy: item.DecisionBy,
                colDecisionDate: item.DecisionDate != null || undefined
                    ? new Date(item.DecisionDate.toString()).toLocaleString() : "",
                colDecisionComments: stringIsNullOrEmpty(item.DecisionComments) ? "" : item.DecisionComments,
                colExistingItemID: item.ExistingItemID,
                colRequestedAction: item.RequestedAction,
                isItemLoaded: true,
                viewScreenControl: item.ApprovalStatus != 'Waiting for Approval' ? true : false,
            });

            //for Modification request, if collection is corporate collection: disabled, else enabled 
            if (item.RequestedAction === 'Modification') {

                //if collection is public, and user requested to mark it as private collection
                if (item.PublicCollection === 0) {
                    //then disable corporate collection flag
                    this.setState({ isCorporateCollection: true });
                }

                //this.setState({ isCorporateCollection: item.CorporateCollection === 1 ? true : false });
            }
            //for Deletion request, user will not mark collection as corporate collection
            if (item.RequestedAction === 'Deletion') {
                this.setState({ isCorporateCollection: true });
                sp.web.lists.getByTitle(this.lstCorpCollectionQueue)
                    .items.filter(`ExistingCollectionID eq ${this.state.colExistingItemID}`).get().then(result => {
                        if (result.length > 0) {
                            if (result[0].Status === 'In Progress') {
                                this.setState({ colRequestShowMessageBar: true, disableColButtons: true });
                            } else {
                                this.setState({ colRequestShowMessageBar: false, disableColButtons: false });
                            }
                        }
                    });

            }
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_openColRequestPanel");
        }

    }

    //to close panel
    private _hideColRequestPanel = (): void => {
        try {
            if (this.isDataRefresh) {
                //refresh detailed data
                this.getRequestData().then((): void => {
                    //clear controls state
                    this.clearColReuqestControlsState();
                });
            } else {
                //clear controls state
                this.clearColReuqestControlsState();
            }

            this.setState({ colRequestIsPanelOpen: false });
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_hideColRequestPanel");
        }

    }

    //Populate panel footer content with "Save" and "Cancel" buttons
    private _onRenderColRequestFooterContent = (): JSX.Element => {
        return (
            <React.Fragment>

                {this.state.colRequestShowMessageBar && (
                    <div>
                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={true}
                            onDismiss={this._closeColRequestMessageBar.bind(this)}
                            dismissButtonAriaLabel="Close"
                        >
                            {this.state.resourceListItems['review_request_corp_collection_delete']}
                        </MessageBar>
                        <br />
                    </div>
                )}
                {this.state.isItemLoaded && !this.state.viewScreenControl && !this.state.disableColButtons && (

                    <div>
                        <DefaultButton
                            className={`${styles.approveBtn}`}
                            text={this.state.resourceListItems['btn_approve_text']}
                            iconProps={{ iconName: 'Checkmark' }}
                            onClick={this._saveColRequest.bind(this, "Approved")}
                        />
                        &nbsp;
                        <DefaultButton
                            className={`${styles.rejectBtn}`}
                            text={this.state.resourceListItems['btn_reject_text']}
                            iconProps={{ iconName: 'Clear' }}
                            onClick={this._saveColRequest.bind(this, "Rejected")}
                        />
                        <br />
                    </div>
                )}
            </React.Fragment >
        );
    }

    // User clicks on save button, new item will create in "ApplicationRequests" list
    private _saveColRequest = async (status: string) => {
        try {
            //console.log(status + ", " + this.currentUserEmail);
            let isFormValid = this.validateColReuqestControls(status);
            //console.log(isFormValid);
            if (isFormValid) {
                //this.setState({ tileRequestShowMessageBar: false });
                let itemCollectionRequest: ICollectionRequest;

                if (status === "Approved") {
                    itemCollectionRequest = {
                        Title: this.state.colTitle,
                        Description: this.state.colDescription,
                        //DefaultMyCollection: this.state.colDefaultMyCollection,
                        UnDeletable: this.state.colUnDeletable,
                        CorporateCollection: this.state.colCorporateCollection,
                        DecisionBy: this.currentUserEmail,
                        DecisionDate: new Date(),
                        DecisionComments: this.state.colDecisionComments,
                        ApprovalStatus: status,
                    };
                } else {
                    itemCollectionRequest = {
                        DecisionBy: this.currentUserEmail,
                        DecisionDate: new Date(),
                        DecisionComments: this.state.colDecisionComments,
                        ApprovalStatus: status,
                    };
                }


                //console.log(itemCollectionRequest);
                //Update existing request data to ApplcationRequest List
                try {
                    sp.web.lists.getByTitle(this.lstCollectionRequest)
                        .items.getById(this.state.colID).update(itemCollectionRequest).then(r => {
                            //console.log(r);
                        });

                    //if collection is publicand admin approved modification request
                    if (this.state.colRequestedAction === 'Modification' && status === "Approved") {
                        //checking if collection entry does not exist in corpcollectionqueue list
                        sp.web.lists.getByTitle(this.lstCorpCollectionQueue)
                            .items.filter(`ExistingCollectionID eq ${this.state.colExistingItemID}`)
                            .get().then(result => {
                                //if no records available in corpcollectionqueue list
                                if (result.length === 0) {
                                    //if admin approves the collection as corporate collection, then create new entry in list
                                    if (this.state.colCorporateCollection === 1) {
                                        //if corporate collection not found in the list, create new entry
                                        let corpCol: ICorpCollectionQueue = {
                                            Title: this.state.colTitle,
                                            ExistingCollectionID: this.state.colExistingItemID,
                                        };
                                        sp.web.lists.getByTitle(this.lstCorpCollectionQueue).items.add(corpCol);
                                    }
                                }
                                //if records found in list
                                else {
                                    //if corporate collection already exists and Admin marked collection as non-corporate
                                    if (this.state.colCorporateCollection === 0) {
                                        //delete item from corp collection list only when it is not 'In Progress'
                                        if (result[0].Status !== 'In Progress') {
                                            //delete item from corp collection list
                                            try {
                                                sp.web.lists
                                                    .getByTitle(this.lstCorpCollectionQueue)
                                                    .items.getById(result[0].Id).delete();
                                            } catch (err) {
                                                this.errorLogging.logError(this.errTitle, this.errModule, "", err, "_saveColRequest Delete corp collection");
                                            }
                                        }
                                    }
                                }
                            });

                    }
                    //if collection is public and admin approved deletion request
                    if (this.state.colRequestedAction === 'Deletion' && status === "Approved") {
                        //hide messagebar
                        this.setState({ colRequestShowMessageBar: false });

                        //delete rerecord from corpcollectionqueue list
                        sp.web.lists
                            .getByTitle(this.lstCorpCollectionQueue)
                            .items
                            .filter(`ExistingCollectionID eq ${this.state.colExistingItemID}`)
                            .get().then(i => {
                                if (i.length > 0) {
                                    //delete item from corp collection list
                                    try {
                                        sp.web.lists
                                            .getByTitle(this.lstCorpCollectionQueue)
                                            .items.getById(i[0].Id).delete();
                                    } catch (err) {
                                        this.errorLogging.logError(this.errTitle, this.errModule, "", err, "_saveColRequest Delete corp collection");
                                    }
                                }
                            });
                    }
                    this.props.callBackForRequestSection("Request reviewed sucessfully.");
                    this.isDataRefresh = true;
                    this._hideColRequestPanel();

                } catch (error) {
                    this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_saveColRequest");
                }

            } else {
                this.ctrColTitle.current.focus();
                this.ctrColDescription.current.focus();
                this.ctrColDecisionComments.current.focus();
                //this.setState({ colRequestShowMessageBar: true });
            }
        } catch (error) {
            this.errorLogging.logError(this.errTitle, this.errModule, "", error, "_saveColRequest");
        }
    }

    //close MessageBar
    private _closeColRequestMessageBar = () => {
        this.setState({ colRequestShowMessageBar: false });
    }

    //validate controls state
    private validateColReuqestControls = (status: string): boolean => {
        let isValid = true;
        if (status === "Approved") {
            if (this.getColRequestTitleErrorMessage(this.state.colTitle) !== "") isValid = false;
            if (this.getColRequestDescriptionErrorMessage(this.state.colDescription) !== "") isValid = false;
            if (this.getColRequestDecisionCommentsErrorMessage(this.state.colDecisionComments) !== "") isValid = false;
        } else {
            if (this.getColRequestDecisionCommentsErrorMessage(this.state.colDecisionComments) !== "") isValid = false;
        }

        return isValid;
    }

    //clear controls state after panel close
    private clearColReuqestControlsState = (): void => {
        this.setState({
            //colRequestIsPanelOpen: false,
            colID: 0,
            colTitle: '',
            //colDefaultMyCollection: 0,
            colDescription: '',
            colPublicCollection: 0,
            colCorporateCollection: 0,
            colStandardOrder: 0,
            colCollectionOwnerId: 0,
            colUnDeletable: 0,
            colRequestedDate: '',
            colApprovalStatus: '',
            colDecisionDate: '',
            colDecisionBy: '',
            colDecisionComments: '',
            colExistingItemID: 0,
            colRequestedAction: '',
            colSelectedTiles: [],
            isItemLoaded: false,
            viewScreenControl: false,
            //colRequestIsPanelOpen: false,
            isCorporateCollection: false,

            colRequestShowMessageBar: false,
            disableColButtons: false,

        });
        this.selectedTiles = [];
        this.isDataRefresh = false;


    }

    //#endregion

    //#region validation
    //Validate Title text
    private getColRequestTitleErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message'] : value.length > 255
                ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }

    //Validate Description text
    private getColRequestDescriptionErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message'] : value.length > 63999
                ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }

    //Validate Decision Comments text
    private getColRequestDecisionCommentsErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value)
            ? this.state.resourceListItems['required_field_validation_message'] : value.length > 63999
                ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //#endregion

    //#region Text Capture
    //capture title
    private _onColRequestTitleChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ colTitle: newValue });
    }

    //capture Description
    private _onColRequestDescriptionChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ colDescription: newValue });
    }

    //Capture Default My Collection text
    // private _onColRequestDefaultMyCollectionChange = (
    //     event: React.MouseEvent<HTMLElement>,
    //     newText: boolean
    // ) => {
    //     this.setState({ colDefaultMyCollection: newText ? 1 : 0 });
    // }

    //capture Public Collection text
    // private _onColRequestPublicCollectionChange = (
    //     event: React.MouseEvent<HTMLElement>,
    //     newText: boolean
    // ) => {
    //     this.setState({ colPublicCollection: newText ? 1 : 0 });
    // }

    //capture Corporate Collection text
    private _onColRequestCorporateCollectionChange = (
        event: React.MouseEvent<HTMLElement>,
        newText: boolean
    ) => {
        this.setState({ colCorporateCollection: newText ? 1 : 0 });
    }

    //capture Undeletable text
    private _onColRequestUndeletableChange = (
        event: React.MouseEvent<HTMLElement>,
        newText: boolean
    ) => {
        this.setState({ colUnDeletable: newText ? 1 : 0 });
    }

    //capture Decision Comments
    private _onColRequestDecisionCommentsChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ colDecisionComments: newValue });
    }
    //#endregion

    //#endregion

}