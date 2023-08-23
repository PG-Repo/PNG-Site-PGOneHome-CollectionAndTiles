import * as React from 'react';
import { ICollectionList } from "../Common/ICollectionList";
import styles from './CollectionsLeftNavigation.module.scss';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Panel, PanelType, IPanelProps } from 'office-ui-fabric-react/lib/Panel';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { IRenderFunction, values } from 'office-ui-fabric-react/lib/Utilities';
import { PrimaryButton, DefaultButton, IconButton, ActionButton } from 'office-ui-fabric-react/lib/Button';
import { arrayMove, SortableContainer, SortableElement } from 'react-sortable-hoc';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PnPHelper } from '../PnPHelper/PnPHelper';
import { IApplicationList } from '../Common/IApplicationList';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { SPPermission } from '@microsoft/sp-page-context';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { Slider, Toggle, MessageBar, MessageBarType, Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType, } from "@pnp/spfx-controls-react/lib";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { ITileRequest } from "../Common/ITileRequest";
import { IBreakingNews } from "../Common/IBreakingNews";
import { sp } from '@pnp/sp';
import { IAddTileRequest } from '../Common/IAddTileRequest';
import { IManageBreakingNews } from '../Common/IManageBreakingNews';
import { stringIsNullOrEmpty } from '@pnp/common';
import { IColorMasterList } from '../Common/IColormasterList';
import { gAnalytics } from '../GA/gAnalytics';


export interface ICollectionsLeftNavigationProps {
    callBackHandlerForTopSettings: any;
    myFollowedCollectionsList: ICollectionList[];
    webpartContext: WebPartContext;
    currentActiveCollection: ICollectionList;
    resourceListItems: any[];
    callBackHandlerForTrackRequests: any;
    callBackHandlerForReviewRequests: any;
    currentSectionTitle?: string;
    callBackForLatestFollowedCollections: any;
    gAnalytics?: gAnalytics;
    isSettingMenuExpanded: boolean;
    loadReadonlyUserProfile: boolean;

    /*Added by Sunny */
    richTextDescription?: string;
}

export interface ICollectionsLeftNavigationState extends IAddTileRequest, IManageBreakingNews {
    myFollowedCollectionItems: ICollectionList[];
    showCreateCollectionPanel: boolean;
    valueRequiredErrorMessage: string;
    collectionName: string;
    collectionDescription: string;
    isPublicCollection: boolean;
    showEditFollowedCollectionPanel: boolean;
    publicCollectionList: ICollectionList[];
    showEditFollowedCollectionInstructionInPanel: boolean;
    currentPublicCollectionDetails?: ICollectionList;
    currentActiveCollection: ICollectionList;
    resourceListItems: any[];
    applicationTiles?: IApplicationList[];
    createCollectionPanelLoading: boolean;
    editFollowedCollectionPanelLoading: boolean;
    currentHoverApplicationTile?: IApplicationList;
    showCreateCollectionInstructionInPanel: boolean;
    currentSelectedCollectionsInEditCollectionPanel?: ICollectionList[];
    showRequiredErrorMessageForApplicationCheckBox: boolean;
    showCalloutMessage: boolean;
    isSettingMenuExpanded: boolean;
    loadReadonlyUserProfile: boolean;

    /*added by sunnny */
    saveControlsErrorMessage?: string;
    requiredFieldErrorMessage?: string;
    tNumber?: string;
    tNumberId?: number;
    configValues?:any;
}

export class CollectionsLeftNavigation extends React.Component<ICollectionsLeftNavigationProps, ICollectionsLeftNavigationState> {

    private pnpHelper: PnPHelper;
    private allApplicationTiles: IApplicationList[] = [];
    private allApplicationRequest: ITileRequest[] = [];
    private allPublicCollections: ICollectionList[] = [];
    private selectedApplicationTilesInCreatePanel: IApplicationList[] = [];
    private selectedCollectionInEditCollectionPanel: ICollectionList[] = [];
    private unSelectedCollectionInEditCollectionPanel: ICollectionList[] = [];
    private currentUserTNumber: string = "";
    private isCurrentUserApprover: boolean;

    private isCurrentUserAdmin = this.props.webpartContext.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);


    private valDescription: string = this.props.richTextDescription;
    private ctrTitle: React.RefObject<ITextField>;
    private ctrDescription: React.RefObject<ITextField>;
    private ctrKeywords: React.RefObject<ITextField>;
    private ctrURL: React.RefObject<ITextField>;

    //added by sunny-06June2020
    private ctrNewsTitle: React.RefObject<ITextField>;
    private ctrNewsTitleFontColor: React.RefObject<ITextField>;
    private ctrNewsBackgroundColor: React.RefObject<ITextField>;
    private ctrNewsContentFontColor: React.RefObject<ITextField>;
    private varExpiryDate: Date = new Date();

    private selectedColorId: number = 0;
    private showError: boolean = false;
    private colorOptions: IDropdownOption[] = [{ key: 0, text: 'Select a category' }];
    private lstColorMaster = "ColorMaster";
    private lstTileRequest = "ApplicationRequests";
    private ApplicationMasterListName: string = "ApplicationMaster";
    private errTitle: string = "PNG-Site-PGOneHome-CollectionAndTiles";
    private errModule: string = "CollectionsLeftNavigation.tsx";

    constructor(props: ICollectionsLeftNavigationProps, state: ICollectionsLeftNavigationState) {
        super(props);
        this.state = {
            myFollowedCollectionItems: this.props.myFollowedCollectionsList,
            showCreateCollectionPanel: false,
            valueRequiredErrorMessage: "",
            collectionName: "",
            collectionDescription: "",
            isPublicCollection: false,
            showEditFollowedCollectionPanel: false,
            publicCollectionList: [],
            isSettingMenuExpanded: this.props.isSettingMenuExpanded,
            showEditFollowedCollectionInstructionInPanel: true,
            currentActiveCollection: this.props.currentActiveCollection,
            resourceListItems: this.props.resourceListItems,
            createCollectionPanelLoading: true,
            editFollowedCollectionPanelLoading: true,
            showCreateCollectionInstructionInPanel: true,
            currentPublicCollectionDetails: { Title: '', ID: '', CollectionOwner: '', CollectionOwnerEmail: '', Description: '', PublicCollection: 0, UnDeletable: 0, CollectionOrder: 0 },
            currentHoverApplicationTile: { Title: '', ID: '', Link: '', Description: '', OwnerEmail: '', SearchKeywords: '', ColorCode: { BgColor: '', ForeColor: '', Title: '' } },
            currentSelectedCollectionsInEditCollectionPanel: this.props.myFollowedCollectionsList,
            showRequiredErrorMessageForApplicationCheckBox: false,
            showCalloutMessage: false,
            loadReadonlyUserProfile: this.props.loadReadonlyUserProfile,

            /* Added by Sunny*/
            requiredFieldErrorMessage: "This field is mandatory",
            tNumber: "",
            tNumberId: 0,
            saveControlsErrorMessage: "",
            tileRequestIsPanelOpen: false,
            tileRequestTitle: "",
            tileRequestDescription: "",
            tileRequestKeywords: "",
            tileRequestUrlLink: "",
            tileRequestIsUrlValid: false,
            tileRequestOwnerEmail: "",
            tileRequestRequestedBy: "",
            tileRequestShowMessageBar: false,
            tileRequestColorCode: 1,
            tileRequestAvailableExternal: 1,
            lstColorMaster: [],
            bNewsIsPanelOpen: false,
            bNewsId: 0,
            bNewsTitle: "",
            bNewsTitleFontColor: "",
            bNewsDescription: "",
            bNewsBackgroundColor: "",
            bNewsContentFontColor: "",
            bNewsContentScrollSpeed: 4,
            bNewsExpiryDate: new Date(),
            //bNewsIsValidExpiryDate: false,
            bNewsIsActive: true,
            bNewsShowMessageBar: false,
            
        };
        this.pnpHelper = new PnPHelper(this.props.webpartContext);
        this.ctrTitle = React.createRef();
        this.ctrDescription = React.createRef();
        this.ctrKeywords = React.createRef();
        this.ctrURL = React.createRef();

        //added by sunny-06June2020
        this.ctrNewsTitle = React.createRef();
        this.ctrNewsTitleFontColor = React.createRef();
        this.ctrNewsBackgroundColor = React.createRef();
        this.ctrNewsContentFontColor = React.createRef();
    }

    public componentWillMount() {
        Promise.all([
            this.pnpHelper.checkCurrentUserApprovalPermission(),
            this.pnpHelper.userProps("TNumber"),this.pnpHelper.getConfigMasterListItems()
        ]).then(([isCurrentUserApprover, currentUserTNumber,configValues]) => {
            this.isCurrentUserApprover = isCurrentUserApprover;
            this.currentUserTNumber = currentUserTNumber;
            this.setState({configValues:configValues}); 
        });

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
                console.log(error);
                this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "componentWillMount");
            }
        }
        try {
            //get Tile request data for validaion
            sp.web.lists.getByTitle(this.lstTileRequest)
                .items.select('Id', 'Link')
                .filter(`ApprovalStatus eq 'Waiting for Approval'`)
                .top(5000)
                .get()
                .then((r): void => {
                    if (r.length > 0) {
                        this.allApplicationRequest = r;
                    }
                });
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "", error, "componentWillMount");
        }
    }


    public componentWillReceiveProps(newProps: ICollectionsLeftNavigationProps) {
        this.setState({
            myFollowedCollectionItems: newProps.myFollowedCollectionsList,
            currentActiveCollection: newProps.currentActiveCollection,
            resourceListItems: newProps.resourceListItems,
            currentSelectedCollectionsInEditCollectionPanel: newProps.myFollowedCollectionsList,
            isSettingMenuExpanded: newProps.isSettingMenuExpanded,
            loadReadonlyUserProfile: newProps.loadReadonlyUserProfile,
        });
    }

    //Added by sunny
    public componentDidMount() {
        //Get Breaking news details
        this._getBreakingNewsDetails();
    }

    private SortableItem = SortableElement(({ value }: { value: ICollectionList }) =>
        <div
            //role={"button"}
            title={value.Description}
            //  tabIndex={0}
            className={`${styles.pgOneCollectionItem} ${this.state.currentActiveCollection !== undefined && this.state.currentActiveCollection["ID"] === value.ID ? `${styles.active}` : ""}`}
            onClick={this._handleCollectionClick.bind(this, value, this.state.myFollowedCollectionItems)}>
            <a href="#">
                {this.currentUserTNumber.toLocaleLowerCase()==value.CollectionOwner.toLocaleLowerCase()&&
                <Icon iconName={this.state.resourceListItems["OwnerIcon"]} />
                }
                 {value.Title}
            </a>
        </div>
    );

    private SortableList = SortableContainer(({ items }: { items: ICollectionList[] }) => {
        return (
            <div className={`${styles.pgOneCollections}`}>
                {items.map((value, index) => (
                    <this.SortableItem key={`item-${index}`} index={index} value={value} />
                ))}
            </div>
        );
    });

    private onSortEnd = ({ oldIndex, newIndex }: { oldIndex: number, newIndex: number }) => {
        try {
            //Google Analytics: collection reorder fired
            this.props.gAnalytics.collectionReorderFired();

            this.setState({
                myFollowedCollectionItems: arrayMove(this.state.myFollowedCollectionItems, oldIndex, newIndex),
            });
            this.props.callBackHandlerForTopSettings(this.state.currentActiveCollection, this.state.myFollowedCollectionItems);
            // update collection order in collection matrix list
            this.pnpHelper.updateSortingOrderInUserCollectionMatrixList(this.state.myFollowedCollectionItems);
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Left panel", error, "onSortEnd");
        }
    }


    public render() {
        return (
            <div className={styles.collectionsLeftNavigation}>
                <div className="ms-hiddenMdDown">
                    <this.SortableList items={this.state.myFollowedCollectionItems} distance={5} lockAxis={"y"}
                        lockToContainerEdges={true} lockOffset={["0%", "100%"]} helperClass={styles.sortableContainer} onSortEnd={this.onSortEnd} />
                </div>
                <div className="ms-hiddenLgUp">
                    <this.SortableList items={this.state.myFollowedCollectionItems} pressDelay={300} lockAxis={"y"}
                        lockToContainerEdges={true} lockOffset={["0%", "100%"]} helperClass={styles.sortableContainer} onSortEnd={this.onSortEnd} />
                </div>
                {/* expand and collapse Setting menu */}
                {!this.state.loadReadonlyUserProfile &&
                    <div>
                        <hr />
                        <div className={styles.settingsMenu}>
                            <ActionButton
                                title={this.state.resourceListItems["leftNavigationCollapseMenuText"]}
                                checked={true}
                                onClick={this._collapseExpandSettingsMenu.bind(this)}
                                iconProps={this.state.isSettingMenuExpanded ? { iconName: 'ChevronDown' } : { iconName: 'ChevronRight' }}
                                aria-expanded={this.state.isSettingMenuExpanded}
                            >
                                <span>{this.state.resourceListItems["leftNavigationCollapseMenuText"]}</span>
                            </ActionButton >
                            {this.state.isSettingMenuExpanded &&
                                <ul className={styles.userActionLinks}>
                                    {/* <li><Link title={this.state.resourceListItems["Add_a_tile_label"]} onClick={this._openTileRequestPanel.bind(this)}><Icon iconName="AppIconDefaultAdd" />{this.state.resourceListItems["Add_a_tile_label"]}</Link></li> */}
                                    <li>
                                        <Link title={this.state.resourceListItems["SnowForm_Label"]}
                                            target="_blank"
                                            data-interception="off"
                                            href={this.state.configValues["SnowForm_Url"]}>
                                            <Icon iconName="AppIconDefaultAdd" />{this.state.resourceListItems["SnowForm_Label"]}</Link>
                                    </li>
                                    
                                    <li className={`${this.props.currentSectionTitle !== undefined && this.props.currentSectionTitle === "Manage my tiles" ? `${styles.active}` : ""}`}>
                                        <Link title={this.state.resourceListItems["Manage_my_tiles_label"]} onClick={this._showManageMyTiles.bind(this, "Manage my tiles")}>
                                            <Icon iconName="TaskManagerMirrored" />
                                            {this.state.resourceListItems["Manage_my_tiles_label"]}</Link>
                                    </li>
                                    <li><Link title={this.state.resourceListItems["create_collection"]} onClick={this._openCreateCollectionPanel.bind(this)}><Icon iconName="AddTo" />{this.state.resourceListItems["create_collection"]}</Link></li>
                                    {/* <li className={`${this.props.currentSectionTitle !== undefined && this.props.currentSectionTitle === "Track requests" ? `${styles.active}` : ""}`}>
                                        <Link title={this.state.resourceListItems["Track_requests_label"]} onClick={this._showTrackRequests.bind(this, "Track requests")}>
                                            <Icon iconName="Trackers" />
                                            {this.state.resourceListItems["Track_requests_label"]}</Link>
                                    </li> */}
                                    {/* <li className={`${this.props.currentSectionTitle !== undefined && this.props.currentSectionTitle === "Track requests" ? `${styles.active}` : ""}`}><Link title={this.state.resourceListItems["Track_requests_label"]} onClick={this._showTrackRequests.bind(this, "Track requests")}><Icon iconName="Trackers" />{this.state.resourceListItems["Track_requests_label"]}</Link></li> */}
                                    <li>
                                        <Link title={this.state.resourceListItems["Track_requests_label"]} 
                                        //onClick={this._showTrackRequests.bind(this, "Track requests")}
                                        href={this.state.configValues["TrackMyRequest_URL"]}
                                        target="_blank"
                                        data-interception="off"
                                        >
                                            <Icon iconName="Trackers" />
                                            {this.state.resourceListItems["Track_requests_label"]}</Link>
                                    </li>
                                    <li>
                                        <Link title={this.state.resourceListItems["edit_collection_follow"]} onClick={this._openEditFollowedCollectionPanel.bind(this)}>
                                            <Icon iconName="Edit" />
                                            {this.state.resourceListItems["edit_collection_follow"]}
                                        </Link>
                                    </li>
                                    <li>
                                        <Link title={this.state.resourceListItems["refresh_page"]} onClick={this._refreshPage.bind(this)}>
                                            <Icon iconName="Refresh" />
                                            {this.state.resourceListItems["refresh_page"]}
                                        </Link>
                                    </li>
                                    {(this.isCurrentUserApprover || this.isCurrentUserAdmin) && // show admin section menu
                                        <div>
                                            <hr />
                                            <li className={`${this.props.currentSectionTitle !== undefined && this.props.currentSectionTitle === "SNOW requests" ? `${styles.active}` : ""}`}>
                                                <Link title={this.state.resourceListItems["Fullfill_ServiceNow_Request"]} onClick={this._showReviewRequests.bind(this, "SNOW requests")}>
                                                    <Icon iconName="CheckList" />
                                                    {this.state.resourceListItems["Fullfill_ServiceNow_Request"]}
                                                </Link>
                                            </li>
                                            {/* <li className={`${this.props.currentSectionTitle !== undefined && this.props.currentSectionTitle === "Review requests" ? `${styles.active}` : ""}`}>
                                                <Link title={this.state.resourceListItems["Review_requests_label"]} onClick={this._showReviewRequests.bind(this, "Review requests")}>
                                                    <Icon iconName="CheckList" />
                                                    {this.state.resourceListItems["Review_requests_label"]}
                                                </Link>
                                            </li> */}
                                            <li>
                                                <Link title={this.state.resourceListItems["Manage_breaking_news"]} onClick={this._openBreakingNewsPanel.bind(this)}>
                                                    <Icon iconName="TaskManagerMirrored" />
                                                    {this.state.resourceListItems["Manage_breaking_news"]}
                                                </Link>
                                            </li>
                                            <li>
                                                <Link
                                                    target="_blank"
                                                    data-interception="off"
                                                    title={this.state.resourceListItems["leftReportDashboardMenuText"]} href={this.props.webpartContext.pageContext.site.absoluteUrl + "/SitePages/AdminDashboard.aspx"}><Icon iconName="ReportDocument" /><span>{this.state.resourceListItems["leftReportDashboardMenuText"]}</span></Link>
                                            </li>
                                        </div>
                                    }
                                </ul>
                            }
                        </div>
                    </div>
                }
                {/* Create Collection Panel */}
                <Panel
                    className={styles.createCollectionPanel}
                    isOpen={this.state.showCreateCollectionPanel}
                    onDismiss={this._hidePanel}
                    type={PanelType.large}
                    closeButtonAriaLabel={this.state.resourceListItems["close"]}
                    headerText={this.state.resourceListItems["create_new_collection_header"]}
                    onRenderFooterContent={this._onRenderFooterContentOnCreateCollection}
                    headerClassName={styles.headerText}
                >
                    <div className={`ms-Grid  ${styles.createCollectionPanelContent}`}>
                        {this.state.createCollectionPanelLoading
                            ? // show loading.. until API success
                            <div className={`ms-Grid-row`}>
                                <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.spinner}`}>
                                    <Spinner size={SpinnerSize.medium} label="Loading..." ariaLive="assertive" labelPosition="right" />
                                </div>
                            </div>
                            :
                            <div className={`ms-Grid-row`}>
                                <div className={`ms-Grid-row`}>
                                    <TextField required label={this.state.resourceListItems["new_collection_name"]} placeholder={"Collection name"} onChange={this._collectionNameChange.bind(this)} maxLength={255} errorMessage={this.state.collectionName.trim() === "" ? this.state.valueRequiredErrorMessage : ""} />
                                    {/* <Checkbox className={styles.chkBoxCollectionType} label={this.state.resourceListItems["new_collection_make_public"]} checked={this.state.isPublicCollection} onChange={this.isPublicCollectionCheckBoxChange.bind(this)} /> */}
                                    <TextField required label={this.state.resourceListItems["new_collection_description"]} placeholder={"Collection description"} onChange={this._collectionDescriptionChange.bind(this)} multiline rows={3} maxLength={2000} errorMessage={this.state.collectionDescription.trim() === "" ? this.state.valueRequiredErrorMessage : ""} />
                                    <Label required>{this.state.resourceListItems["create_collection_select_tile_label"]}</Label>

                                </div>
                                <div className={`ms-Grid-row`}>
                                    <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg4 ${styles.createCollectionsLeftPart}`}>
                                        <SearchBox
                                            placeholder="Search a tile"
                                            onChange={this._searchApplicationTilesList.bind(this)}
                                            onSearch={this._searchApplicationTilesList.bind(this)} />

                                        {(this.state.showRequiredErrorMessageForApplicationCheckBox && this.selectedApplicationTilesInCreatePanel.length == 0) && <Label className={styles.requiredCheckBoxCheck}>{this.state.resourceListItems["create_collection_select_tile_error_message"]}</Label>}
                                        <ul className={styles.applicationTilesList} aria-label={this.state.resourceListItems["create_collection_select_tile_label"]}>
                                            {
                                                this.state.applicationTiles.map((applicationTile: any, i: any) =>
                                                    <li title={applicationTile.Title} aria-live="polite" onFocus={this._hoverApplicationTileList.bind(this, applicationTile)} onMouseOver={this._hoverApplicationTileList.bind(this, applicationTile)}>
                                                        <Checkbox
                                                            label={applicationTile.Title}
                                                            onChange={(event) => this._handleApplicationTileSelection(event, applicationTile)}
                                                        />
                                                    </li>
                                                )
                                            }
                                        </ul>
                                        {this.allApplicationTiles.length != this.state.applicationTiles.length
                                            &&
                                            <div aria-live="polite" className="sr-only" role="status">{this.state.applicationTiles.length} suggestions found</div>
                                        }
                                    </div>
                                    <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg8 ${styles.createCollectionsRightPart}`}>
                                        <div className={this.state.showCreateCollectionInstructionInPanel ? styles.showCreateCollectionInstructionInPanel : styles.hideCreateCollectionInstructionInPanel}>
                                            <p>{this.state.resourceListItems["create_collection_tile_instruction_line1"]}</p>
                                            <p>{this.state.resourceListItems["create_collection_tile_instruction_line2"]}</p>
                                        </div>
                                        <div className={this.state.showCreateCollectionInstructionInPanel ? styles.hideCreateCollectionInstructionInPanel : styles.showCreateCollectionInstructionInPanel}>
                                            <div className={styles.tileColor} title={"Tile Category: " + this.state.currentHoverApplicationTile.ColorCode.Title} style={{ backgroundColor: this.state.currentHoverApplicationTile.ColorCode.BgColor }} />
                                            <h2 title={this.state.currentHoverApplicationTile.Title} className={styles.applicationTileTitle}>{this.state.currentHoverApplicationTile.Title}</h2>
                                            <p className={styles.applicationTileDescription} title={this.state.currentHoverApplicationTile.Description}>{this.state.currentHoverApplicationTile.Description}</p>
                                            <Link className={styles.gotoSiteLink}
                                                title={this.state.resourceListItems["goToSite"]}
                                                target="_blank"
                                                href={this.state.currentHoverApplicationTile.Link}
                                                onClick={() => {
                                                    //Google Analytics: tile link visited from list (Create a new collection)
                                                    this.props.gAnalytics.appHitFromList(this.state.currentHoverApplicationTile.Link, this.state.currentHoverApplicationTile.ID);
                                                    //window.open(this.state.currentHoverApplicationTile.Link, '_blank');
                                                }}
                                                data-interception="off">{this.state.resourceListItems["goToSite"]}</Link>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        }
                    </div>
                </Panel>

                {/* Edit Collection which you follow */}
                <Panel
                    className={styles.editFollowedCollectionPanel}
                    isOpen={this.state.showEditFollowedCollectionPanel}
                    onDismiss={this._hidePanel}
                    type={PanelType.large}
                    closeButtonAriaLabel={this.state.resourceListItems["close"]}
                    headerText={this.state.resourceListItems["edit_collection_follow"]}
                    onRenderFooterContent={this._onRenderFooterContentEditFollowedCollection}
                    headerClassName={styles.headerText}
                >
                    <div className={`ms-hiddenLgUp ${styles.infoIconArea}`}>
                        <Icon iconName="Info"
                            onClick={() => { this.setState({ showCalloutMessage: !this.state.showCalloutMessage }); }}
                        />
                    </div>
                    {this.state.showCalloutMessage &&
                        <Callout
                            role="alertdialog"
                            gapSpace={0}
                            setInitialFocus
                            target={`.${styles.infoIconArea}`}
                            className={styles.callout}
                            onDismiss={() => { this.setState({ showCalloutMessage: !this.state.showCalloutMessage }); }}
                        >
                            <div className={styles.editFollowedCollectionCallOutContent}>
                                <p>{this.state.resourceListItems["edit_collection_follow_instruction_line1_for_mobile"]}</p>
                                <p>{this.state.resourceListItems["edit_collection_follow_instruction_line2_for_mobile"]}</p>
                            </div>
                        </Callout>
                    }
                    <div className={`ms-Grid ${styles.editFollowedCollectionContent}`}>
                        {this.state.editFollowedCollectionPanelLoading
                            ? <div className={`ms-Grid-row`}>
                                <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.spinner}`}>
                                    <Spinner size={SpinnerSize.medium} label="Loading..." ariaLive="assertive" labelPosition="right" />
                                </div>
                            </div>
                            :
                            <div className={`ms-Grid-row`}>
                                <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg4 ${styles.editFollowedCollectionLeftPart}`}>
                                    <SearchBox
                                        placeholder="Search Collections"
                                        onChange={this._searchPublicCollectionList.bind(this)}
                                        onSearch={this._searchPublicCollectionList.bind(this)} />
                                    <ul className={styles.collectionItems} aria-label={this.state.resourceListItems["create_collection_select_tile_label"]}>
                                        {
                                            this.state.publicCollectionList.map((publicCollection: any, i: any) =>
                                                <li title={publicCollection.Title} aria-live="polite" onFocus={this._hoverPublicCollectionList.bind(this, publicCollection)} onMouseOver={this._hoverPublicCollectionList.bind(this, publicCollection)}>
                                                    <Checkbox
                                                        label={publicCollection.Title}
                                                        checked={
                                                            this.state.currentSelectedCollectionsInEditCollectionPanel.some((followedCollection: any) => {
                                                                return followedCollection.ID === publicCollection.ID;
                                                            })
                                                        }
                                                        onChange={(event) => this._handleCollectionSelection(event, publicCollection)}
                                                        disabled={String(publicCollection.CollectionOwner).toLocaleLowerCase() == this.currentUserTNumber.toLocaleLowerCase()}
                                                    />
                                                </li>
                                            )
                                        }
                                    </ul>
                                </div>
                                <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg8 ${styles.editFollowedCollectionRightPart}`}>
                                    <div className={`ms-hiddenMdDown ${this.state.showEditFollowedCollectionInstructionInPanel ? `${styles.showEditFollowedCollectionInstructionInPanel}` : `${styles.hideEditFollowedCollectionInstructionInPanel}`} `}>
                                        <p>{this.state.resourceListItems["edit_collection_follow_instruction_line1_for_desktop"]}</p>
                                        <p>{this.state.resourceListItems["edit_collection_follow_instruction_line2_for_desktop"]}</p>
                                    </div>
                                    <div className={this.state.showEditFollowedCollectionInstructionInPanel ? styles.hideEditFollowedCollectionInstructionInPanel : styles.showEditFollowedCollectionInstructionInPanel}>
                                        <h2 className={styles.collectionTitle} title={this.state.currentPublicCollectionDetails.Title}>{this.state.currentPublicCollectionDetails.Title}</h2>
                                        <p className={styles.collectionDescription} title={this.state.currentPublicCollectionDetails.Description}>{this.state.currentPublicCollectionDetails.Description}</p>
                                        <div className={styles.collectionDetailsLink}>
                                            <Link
                                                title={this.state.resourceListItems["goToCollection"]}
                                                data-interception="off"
                                                onClick={() => {
                                                    let url = this.props.webpartContext.pageContext.site.absoluteUrl + "?collectionId=" + this.state.currentPublicCollectionDetails.ID;
                                                    //Google Analytics: Collection Hit from List (Edit which collection you follow) 
                                                    this.props.gAnalytics.collectionHitFromList(this.state.currentPublicCollectionDetails.Title, this.state.currentPublicCollectionDetails.ID);
                                                    window.open(url, "_blank");
                                                }}
                                            >{this.state.resourceListItems["goToCollection"]}</Link>
                                            <Link title={this.state.resourceListItems["emailOwner"]} href={"mailto:" + this.state.currentPublicCollectionDetails.CollectionOwnerEmail}>{this.state.resourceListItems["emailOwner"]}</Link>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        }
                    </div>
                </Panel>

                {/*Tile Request Panel => Start*/}
                <Panel
                    isFooterAtBottom={true}
                    isOpen={this.state.tileRequestIsPanelOpen}
                    onDismiss={this._hideTileRequestPanel}
                    type={PanelType.smallFixedFar}
                    closeButtonAriaLabel={this.state.resourceListItems["close"]}
                    headerText={this.state.resourceListItems["add_tile_header_label"]}
                    onRenderFooterContent={this._onRenderTileRequestFooterContent}
                    isLightDismiss={false}
                    className={styles.tileRequestAddPanel}
                    headerClassName={styles.headerText}
                    style={{ fontWeight: "bold" }}
                >
                    <Label className={styles["ms-font-m"]}>{this.state.resourceListItems["Add_Tile_Message"]}</Label>
                    <br></br>
                    <div className={`ms-Grid  ${styles.tileRequestPanelContent}`}>
                        <TextField
                            label={this.state.resourceListItems["add_tile_title_label"]}
                            ariaLabel={this.state.resourceListItems["add_tile_title_label"]}
                            placeholder="Tile name"
                            onChange={this._onTileRequestTitleChange.bind(this)}
                            required={true}
                            //errorMessage={this.state.valueRequiredErrorMessage}
                            onGetErrorMessage={this.getTileRequestTitleErrorMessage.bind(this)}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrTitle}
                            maxLength={255}
                        />
                        <TextField
                            label={this.state.resourceListItems["add_tile_description_label"]}
                            ariaLabel={this.state.resourceListItems["add_tile_description_label"]}
                            placeholder="Tile description"
                            onChange={this._onTileRequestDescriptionChange.bind(this)}
                            required={true}
                            multiline
                            rows={3}
                            //errorMessage={this.state.valueRequiredErrorMessage}
                            onGetErrorMessage={this.getTileRequestDescriptionErrorMessage.bind(
                                this
                            )}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrDescription}
                        />
                        <TextField
                            label={this.state.resourceListItems["add_tile_keywords_label"]}
                            ariaLabel={this.state.resourceListItems["add_tile_keywords_label"]}
                            placeholder="Comma separated keywords. e.g. Home, Document"
                            onChange={this._onTileRequestKeywordsChange.bind(this)}
                            required={true}
                            multiline
                            rows={3}
                            //errorMessage={this.state.valueRequiredErrorMessage}
                            onGetErrorMessage={this.getTileRequestKeywordsErrorMessage.bind(
                                this
                            )}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrKeywords}
                        />
                        <TextField
                            label={this.state.resourceListItems["add_tile_url_link_label"]}
                            ariaLabel={this.state.resourceListItems["add_tile_url_link_label"]}
                            placeholder="e.g. https://pgone.pg.com"
                            onChange={this._onTileRequestURLLinkChange.bind(this)}
                            required={true}
                            multiline
                            rows={3}
                            //errorMessage={this.state.valueRequiredErrorMessage}
                            onGetErrorMessage={this.getTileRequestURLLinkErrorMessage.bind(
                                this
                            )}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrURL}
                        />
                        <Link
                            className={styles.linkText}
                            href={this.state.tileRequestUrlLink}
                            target="_blank"
                            data-interception="off"
                            disabled={this.state.tileRequestIsUrlValid ? false : true}
                        >
                            {this.state.resourceListItems["add_tile_try_link_label"]}
                        </Link>
                        <Dropdown
                            options={this.colorOptions}
                            selectedKey={this.selectedColorId}
                            label={this.state.resourceListItems["add_tile_category_label"]}
                            ariaLabel={this.state.resourceListItems["add_tile_category_label"]}
                            required={true}
                            onChange={this._onColorCodeChange.bind(this)}
                            data-is-focusable={true}
                            errorMessage={this.selectedColorId === 0 && this.showError && this.state.resourceListItems['required_field_validation_message']}

                        ></Dropdown>
                        <Toggle
                            className={styles.isAvailableExternalLabel}
                            label={this.state.resourceListItems['add_tile_available_external_label']}
                            onText="Yes"
                            offText="No"
                            checked={this.state.tileRequestAvailableExternal == 1 ? false : true}
                            onChange={this._onTileRequestAvailableExternalChange.bind(this)}
                        />
                        {/* <TextField
                            label="Owner Email"
                            placeholder="e.g. email@pg.com"
                            value={this.state.tileRequestOwnerEmail}
                            validateOnLoad={false}
                            validateOnFocusOut={true}
                            onChange={this._onTileRequestOwnerEmailChange.bind(this)}
                            errorMessage={this.state.valueRequiredErrorMessage}
                            onGetErrorMessage={this.getTileRequestOwnerEmailErrorMessage.bind(this)}
                        /> */}
                    </div>
                </Panel>
                {/*Tile Request Panel => End*/}

                {/*Manage Breaking News Panel => Start*/}
                <Panel
                    isFooterAtBottom={true}
                    isOpen={this.state.bNewsIsPanelOpen}
                    onDismiss={this._hideBreakingNewsPanel}
                    type={PanelType.smallFixedFar}
                    closeButtonAriaLabel={this.state.resourceListItems["close"]}
                    headerText={this.state.resourceListItems["bnews_header_title_label"]}
                    onRenderFooterContent={this._onRenderBreakingNewsFooterContent}
                    isLightDismiss={false}
                    className={styles.breakingNewsPanel}
                    headerClassName={styles.headerText}
                    style={{ fontWeight: "bold" }}
                >
                    <div className={`ms-Grid  ${styles.tileRequestPanelContent}`}>
                        <br />
                        <TextField
                            label={this.state.resourceListItems["bnews_title_label"]}
                            ariaLabel={this.state.resourceListItems["bnews_title_label"]}
                            placeholder="Enter breaking news title"
                            onChange={this._onBreakingNewsTitleChange.bind(this)}
                            required={true}
                            onGetErrorMessage={this.getBreakingNewsTitleErrorMessage.bind(
                                this
                            )}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrNewsTitle}
                            value={this.state.bNewsTitle}
                        />
                        <TextField
                            label={this.state.resourceListItems["bnews_title_fontcolor_label"]}
                            ariaLabel={this.state.resourceListItems["bnews_title_fontcolor_label"]}
                            placeholder="Hex color code. e.g. #008AFB"
                            onChange={this._onBreakingNewsTitleFontColorChange.bind(this)}
                            required={true}
                            onGetErrorMessage={this.getBreakingNewsTitleFontColorErrorMessage.bind(
                                this
                            )}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrNewsTitleFontColor}
                            value={this.state.bNewsTitleFontColor}
                        />
                        <label>
                            {this.state.resourceListItems["bnews_description_label"]}{" "}
                            <span style={{ color: "#B66D52" }}>*</span>
                        </label>

                        <RichText
                            isEditMode={true}
                            value={this.state.bNewsDescription}
                            onChange={(text) => this._onBreakingNewsDescriptionChange(text)}
                            placeholder="Enter breaking news description"
                            styleOptions={{
                                showBold: true,
                                showItalic: true,
                                showUnderline: true,
                                showLink: true,
                                showAlign: false,
                                showList: false,
                                showMore: false,
                                showStyles: false,
                            }}
                            className="rte--read"
                        ></RichText>

                        <TextField
                            label={this.state.resourceListItems["bnews_background_color_label"]}
                            ariaLabel={this.state.resourceListItems["bnews_background_color_label"]}
                            placeholder="Hex color code. e.g. #012169"
                            onChange={this._onBreakingNewsBgColorChange.bind(this)}
                            required={true}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrNewsBackgroundColor}
                            value={this.state.bNewsBackgroundColor}
                            onGetErrorMessage={this.getBreakingNewsBgColorErrorMessage.bind(
                                this
                            )}
                        />
                        <TextField
                            label={this.state.resourceListItems["bnews_content_fontcolor_label"]}
                            ariaLabel={this.state.resourceListItems["bnews_content_fontcolor_label"]}
                            placeholder="Hex color code. e.g. #008AFB"
                            onChange={this._onBreakingNewsFontColorChange.bind(this)}
                            required={true}
                            onGetErrorMessage={this.getBreakingNewsFontColorErrorMessage.bind(
                                this
                            )}
                            validateOnLoad={false}
                            validateOnFocusIn={true}
                            validateOnFocusOut={true}
                            componentRef={this.ctrNewsContentFontColor}
                            value={this.state.bNewsContentFontColor}
                        />
                        <DateTimePicker
                            label={this.state.resourceListItems["bnews_expiration_date_label"]}
                            placeholder="Expiration date and time"
                            dateConvention={DateConvention.DateTime}
                            timeConvention={TimeConvention.Hours12}
                            showSeconds={false}
                            value={this.state.bNewsExpiryDate}
                            onChange={this._onBreakingNewsExpiryDateChange.bind(this)}
                            onGetErrorMessage={this.getBreakingNewsExpiryDateErrorMessage.bind(
                                this
                            )}
                            timeDisplayControlType={TimeDisplayControlType.Dropdown}
                            showLabels={false}
                        />
                        <Toggle
                            label={this.state.resourceListItems["bnews_isactive_label"]}
                            ariaLabel={this.state.resourceListItems["bnews_isactive_label"]}
                            onText="Yes"
                            offText="No"
                            checked={this.state.bNewsIsActive}
                            onChange={this._onBreakingNewsIsActiveChange.bind(this)}
                        />
                        <Slider
                            label={this.state.resourceListItems["bnews_content_scroll_speed"]}
                            ariaLabel={this.state.resourceListItems["bnews_content_scroll_speed"]}
                            min={1}
                            max={20}
                            step={1}
                            defaultValue={4}
                            showValue={true}
                            value={this.state.bNewsContentScrollSpeed}
                            onChange={(value: number) =>
                                this.setState({ bNewsContentScrollSpeed: value })
                            }
                        // snapToStep={true}
                        />
                    </div>
                </Panel>
                {/*Manage Breaking News Panel => End*/}
            </div >
        );
    }

    //Footer buttons for Create Collection Panel
    private _onRenderFooterContentOnCreateCollection = (props: IPanelProps, defaultRender: IRenderFunction<IPanelProps>): JSX.Element => {
        return (
            <React.Fragment>
                <div className={styles.btnContainer}>
                    <PrimaryButton title={this.state.resourceListItems["save"]} onClick={this._saveNewCollectionAndApplications.bind(this)} text={this.state.resourceListItems["save"]} />
                    <DefaultButton title={this.state.resourceListItems["cancel"]} onClick={this._hidePanel.bind(this)} text={this.state.resourceListItems["cancel"]} />
                </div>
            </React.Fragment >
        );
    }

    //Footer buttons for Edit Followed Collection
    private _onRenderFooterContentEditFollowedCollection = (props: IPanelProps, defaultRender: IRenderFunction<IPanelProps>): JSX.Element => {
        return (
            <React.Fragment>
                <div className={styles.btnContainer}>
                    <PrimaryButton title={this.state.resourceListItems["save"]} onClick={this._saveFollowedCollection.bind(this)} text={this.state.resourceListItems["save"]} />
                    <DefaultButton title={this.state.resourceListItems["cancel"]} onClick={this._hidePanel.bind(this)} text={this.state.resourceListItems["cancel"]} />
                </div>
            </React.Fragment>
        );
    }

    // left navigation collapse Expand menu
    private _collapseExpandSettingsMenu() {
        this.setState({
            isSettingMenuExpanded: !this.state.isSettingMenuExpanded,
        });
    }
    //#region Manage My tiles
    // Manage My Tiles
    private _showManageMyTiles = (requestType: string): void => {
        this.props.callBackHandlerForTrackRequests(requestType, this.state.myFollowedCollectionItems);
    }

    // Track request events
    private _showTrackRequests = (requestType: string): void => {
        this.props.callBackHandlerForTrackRequests(requestType, this.state.myFollowedCollectionItems);
    }

    // Review request events
    private _showReviewRequests = (requestType: string): void => {
        this.props.callBackHandlerForReviewRequests(requestType, this.state.myFollowedCollectionItems);
    }

    public _handleCollectionClick(item: any, myFollowedCollectionItems: ICollectionList[]) {
        //Google Analytics: collection onclick
        this.props.gAnalytics.collectionHit(item.Title, item.ID);
        this.props.callBackHandlerForTopSettings(item, myFollowedCollectionItems);
    }
    private _openCreateCollectionPanel = (): void => {
        //Google Analytics: create Collection Called
        this.props.gAnalytics.createCollectionCalled();

        this.setState({
            showCreateCollectionPanel: true,
            createCollectionPanelLoading: true
        });
        this.pnpHelper.getApplicationTiles()
            .then((applicationTiles: IApplicationList[]) => {
                this.allApplicationTiles = applicationTiles;
                this.setState({
                    applicationTiles: applicationTiles,
                    createCollectionPanelLoading: false,
                });
            });


    }

    //Handling Application check and uncheck events
    private _handleApplicationTileSelection(event: any,
        applicationTileCheckBoxValue: any) {
        if (event.target.checked) {
            //Google Analytics: tile checked from list (Create a new collection)
            this.props.gAnalytics.appListAdd(
                applicationTileCheckBoxValue.Title,
                applicationTileCheckBoxValue.Link,
                this.state.currentActiveCollection.Title,
                this.state.currentActiveCollection.ID,
                applicationTileCheckBoxValue.ID);
            this.selectedApplicationTilesInCreatePanel.push(applicationTileCheckBoxValue);
            this.setState({
                showRequiredErrorMessageForApplicationCheckBox: false,
            });
        }
        else if (!event.target.checked) {
            //Google Analytics: tile unchecked from list (Create a new collection)
            this.props.gAnalytics.appListRemove(
                applicationTileCheckBoxValue.Title,
                applicationTileCheckBoxValue.Link,
                this.state.currentActiveCollection.Title,
                this.state.currentActiveCollection.ID,
                applicationTileCheckBoxValue.ID);

            this.selectedApplicationTilesInCreatePanel = this.selectedApplicationTilesInCreatePanel
                .filter((unSelectedApplicationTile: any) => {
                    return unSelectedApplicationTile.ID !== applicationTileCheckBoxValue.ID;
                });
        }
    }
    // hide the panels
    private _hidePanel = (): void => {
        this.setState({
            showCreateCollectionPanel: false,
            showEditFollowedCollectionPanel: false,
            showEditFollowedCollectionInstructionInPanel: true,
            publicCollectionList: this.allPublicCollections,
            showCreateCollectionInstructionInPanel: true,
            applicationTiles: this.allApplicationTiles,
            currentSelectedCollectionsInEditCollectionPanel: this.props.myFollowedCollectionsList,
            showRequiredErrorMessageForApplicationCheckBox: false,
            isPublicCollection: false,
            valueRequiredErrorMessage: ""
        });
    }
    // CollectionName Change event
    private _collectionNameChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ): void => {
        this.setState({ collectionName: newValue });
    }

    //Collection Description Change event
    private _collectionDescriptionChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ): void => {
        this.setState({ collectionDescription: newValue });
    }

    //isPublic Collection Check box change
    private isPublicCollectionCheckBoxChange = (
        event: React.FormEvent<HTMLElement>,
        checked: boolean): void => {
        this.setState({
            isPublicCollection: checked
        });
    }

    //Save New Collection functionalities
    private _saveNewCollectionAndApplications = async () => {
        try {
            //Google Analytics: create Collection save
            this.props.gAnalytics.createCollectionSave();

            //required field validation
            if (this.state.collectionName.trim() === "" || this.state.collectionDescription.trim() === "" || this.selectedApplicationTilesInCreatePanel.length == 0) {
                this.setState({
                    showCreateCollectionPanel: true,
                    valueRequiredErrorMessage: this.state.resourceListItems["required_field_validation_message"],
                    showRequiredErrorMessageForApplicationCheckBox: true,
                });
            }
            else {
                let currentUserItem = await this.pnpHelper.getCurrentUserItemID(this.currentUserTNumber);
                let currentMaximumCollectionOrder = Math.max.apply(Math, this.state.myFollowedCollectionItems.map((followedCollectionItem: any) => { return followedCollectionItem.CollectionOrder; }));
                let currentMyFollowedCollectionsLength = this.state.myFollowedCollectionItems.length;
                let collectionItem: ICollectionList = {
                    Title: this.state.collectionName,
                    Description: this.state.collectionDescription,
                    DefaultMyCollection: 0,
                    PublicCollection: Number(this.state.isPublicCollection),
                    CorporateCollection: 0,
                    StandardOrder: 0,
                    UnDeletable: 0,
                    CollectionOwnerId: currentUserItem[0].ID
                };
                // add private collection to Collection master
                if (!this.state.isPublicCollection) {
                    this.pnpHelper.addPrivateCollectionsToMasterList(collectionItem).then((collectionItemAdded: any) => {
                        let userCollectionItem: any = {
                            CollectionIDId: collectionItemAdded.ID,
                            UserIDId: currentUserItem[0].ID,
                            CollectionOrder: currentMaximumCollectionOrder + 1
                        };
                        Promise.all([
                            this.pnpHelper.addItemToCollectionApplicationMatrixList(this.selectedApplicationTilesInCreatePanel, collectionItemAdded.ID),
                            this.pnpHelper.addItemToUserCollectionMatrixList(userCollectionItem)
                        ]).then(() => {
                            this.props.callBackForLatestFollowedCollections(false, true, currentMyFollowedCollectionsLength);
                        });
                    });
                }
                //add public collection to collection requests
                else {
                    let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentActiveCollection["ID"]);
                    this.pnpHelper.addPublicCollectionsToRequestList(collectionItem).then((collectionRequestItemAdded: any) => {
                        this.pnpHelper.addItemToCollectionApplicationMatrixRequestsList(this.selectedApplicationTilesInCreatePanel, collectionRequestItemAdded.ID)
                            .then(() => {
                                this.props.callBackForLatestFollowedCollections(false, false, currentItemIndexInMyFollowedCollectionList, this.state.resourceListItems["create_public_collection_success_message"]);
                            }
                            );
                    });
                }
                this.setState({ showCreateCollectionPanel: false });
                this.props.callBackForLatestFollowedCollections(true);
            }
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _saveNewCollectionAndApplications()", error, "Create a new Collection");
        }
    }

    // search application lists based on title, description and searchkeywords fields
    private _searchApplicationTilesList = (searchValue: string): void => {
        try {
            //Google Analytics: tile search from list (Create a new collection)
            this.props.gAnalytics.appListSearch(this.state.currentActiveCollection.Title, this.state.currentActiveCollection.ID, searchValue);

            let filteredApplicationTileList: IApplicationList[] = [];
            this.allApplicationTiles.map((data: any, i: any) => {
                if (String(data["Title"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(data["Description"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(data["SearchKeywords"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1) {
                    filteredApplicationTileList.push(data);
                }
            });

            this.setState({
                applicationTiles: filteredApplicationTileList
            });
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _searchApplicationTilesList()", error, "Search Application List");
        }
    }

    // render current application tile and description on hover of application tile
    private _hoverApplicationTileList = (currentHoverApplicationTile: ICollectionList): void => {
        this.setState({
            currentHoverApplicationTile: currentHoverApplicationTile,
            showCreateCollectionInstructionInPanel: false
        });
    }
    //#endregion

    //#region  Edit Followed Collection Panel Methods
    private _searchPublicCollectionList = (searchValue: string): void => {
        try {
            //Google Analytics: collection searched from list (Edit which collections you follow)
            this.props.gAnalytics.collectionListSearch(searchValue);

            let filteredPublicCollectionList: ICollectionList[] = [];
            this.allPublicCollections.map((data: any, i: any) => {
                if (String(data["Title"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(data["Description"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1) {
                    filteredPublicCollectionList.push(data);
                }
            });
            this.setState({
                publicCollectionList: filteredPublicCollectionList
            });
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _searchPublicCollectionList()", error, "Search Public Collection List");
        }
    }

    private _handleCollectionSelection(event: any,
        publicCollectionCheckBoxValue: any) {
        try {
            if (event.target.checked) {
                //Google Analytics: collection checked from list (Edit which collections you follow)
                this.props.gAnalytics.collectionListFollow(publicCollectionCheckBoxValue.Title, publicCollectionCheckBoxValue.ID);

                this.selectedCollectionInEditCollectionPanel.push(publicCollectionCheckBoxValue);
                var pushedCurrentSelectedCollectionToFollowedList = this.state.currentSelectedCollectionsInEditCollectionPanel.concat(this.selectedCollectionInEditCollectionPanel);
                this.setState({
                    // just to remove duplicates before setting the state
                    currentSelectedCollectionsInEditCollectionPanel: pushedCurrentSelectedCollectionToFollowedList.filter((item: any, pos: any, self: any) => {
                        return self.indexOf(item) == pos;
                    })
                });

                // push only newly added records
                this.selectedCollectionInEditCollectionPanel = this.selectedCollectionInEditCollectionPanel.filter((selectedCollectionItem: any) => {
                    return !this.props.myFollowedCollectionsList.some((alreadyFollowedCollectionItem: any) => {
                        return alreadyFollowedCollectionItem.ID === selectedCollectionItem.ID;
                    });
                });

                this.unSelectedCollectionInEditCollectionPanel = this.unSelectedCollectionInEditCollectionPanel
                    .filter((unSelectedFollowedCollection: any) => {
                        return unSelectedFollowedCollection.ID !== publicCollectionCheckBoxValue.ID;
                    });
            }
            else if (!event.target.checked) {
                //Google Analytics: collection unchecked from list (Edit which collections you follow)
                this.props.gAnalytics.collectionListUnfollow(publicCollectionCheckBoxValue.Title, publicCollectionCheckBoxValue.ID);

                this.selectedCollectionInEditCollectionPanel = this.selectedCollectionInEditCollectionPanel
                    .filter((unSelectedFollowedCollection: any) => {
                        return unSelectedFollowedCollection.ID !== publicCollectionCheckBoxValue.ID;
                    });
                this.unSelectedCollectionInEditCollectionPanel.push(publicCollectionCheckBoxValue);
                //// delete only already followed records
                this.unSelectedCollectionInEditCollectionPanel = this.unSelectedCollectionInEditCollectionPanel.filter((unSelectedCollectionItem: any) => {
                    return this.props.myFollowedCollectionsList.some((alreadyFollowedCollectionItem: any) => {
                        return alreadyFollowedCollectionItem.ID === unSelectedCollectionItem.ID;
                    });
                });

                this.unSelectedCollectionInEditCollectionPanel = this.unSelectedCollectionInEditCollectionPanel.map((value: any) => {
                    this.props.myFollowedCollectionsList.map((collectionItem: any) => {
                        if (collectionItem.ID == value.ID) {
                            value["UserCollectionMatrixItemID"] = collectionItem["UserCollectionMatrixItemID"];
                        }
                    });
                    return value;
                });

                this.unSelectedCollectionInEditCollectionPanel = this.unSelectedCollectionInEditCollectionPanel.filter((item: any, pos: any, self: any) => {
                    return self.indexOf(item) == pos;
                });
                this.setState({
                    currentSelectedCollectionsInEditCollectionPanel: this.state.currentSelectedCollectionsInEditCollectionPanel.filter((unSelectedFollowedCollection: any) => {
                        return unSelectedFollowedCollection.ID !== publicCollectionCheckBoxValue.ID;
                    })
                });
            }
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _handleCollectionSelection()", error, "Handle Collection Selection");
        }
    }

    private _saveFollowedCollection = async (): Promise<void> => {
        try {
            //Google Analytics: Edit which collections you follow save clicked
            this.props.gAnalytics.editCollectionsFollowedSaved();

            if (this.selectedCollectionInEditCollectionPanel.length == 0 && this.unSelectedCollectionInEditCollectionPanel.length == 0) {
                this._hidePanel();
            }
            else {
                let currentUserTNumber = await this.pnpHelper.userProps("TNumber");
                let currentUserItem = await this.pnpHelper.getCurrentUserItemID(currentUserTNumber);
                let currentMaximumCollectionOrder = Math.max.apply(Math, this.state.myFollowedCollectionItems.map((followedCollectionItem: any) => { return followedCollectionItem.CollectionOrder; }));
                let currentMyFollowedCollectionsLength = this.state.currentSelectedCollectionsInEditCollectionPanel.length - 1;
                Promise.all([
                    this.pnpHelper.addMulipleItemsToUserCollectionMatrixList(currentUserItem, this.selectedCollectionInEditCollectionPanel, currentMaximumCollectionOrder),
                    this.pnpHelper.deleteItemsOnUserCollectionMatrixList(this.unSelectedCollectionInEditCollectionPanel)
                ]).then(() => {
                    this.props.callBackForLatestFollowedCollections(false, true, currentMyFollowedCollectionsLength);
                });
                this.setState({ showEditFollowedCollectionPanel: false });
                this.props.callBackForLatestFollowedCollections(true);
            }
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _saveFollowedCollection()", error, "Save Followed Collection");
        }
    }

    private _hoverPublicCollectionList = (currentHoverPublicCollection: ICollectionList): void => {
        this.setState({
            currentPublicCollectionDetails: currentHoverPublicCollection,
            showEditFollowedCollectionInstructionInPanel: false
        });
    }

    private _refreshPage = (): void => {
        localStorage.clear();
        location.reload();
    }

    private _openEditFollowedCollectionPanel = (): void => {
        try {
            //Google Analytics: Edit which collections you follow
            this.props.gAnalytics.editCollectionsFollowedClicked();

            this.setState({
                showEditFollowedCollectionPanel: true,
                editFollowedCollectionPanelLoading: true
            });
            this.pnpHelper.getPublicCollections()
                .then((publicCollectionsList: ICollectionList[]) => {
                    this.allPublicCollections = publicCollectionsList;
                    this.setState({
                        publicCollectionList: publicCollectionsList,
                        editFollowedCollectionPanelLoading: false,
                    });
                });
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _openEditFollowedCollectionPanel()", error, "Open Edit Followed Collection Panel");
        }
    }
    //#endregion

    //#region Add a tile News Events and Validation

    //#region Tile Request Other events
    //Called when user clicked on "Add a tile" option from left pane
    private _openTileRequestPanel = async () => {
        try {
            this.setState({ tileRequestIsPanelOpen: true });

            //Getting current user details using @PnP library
            this.setState({ tileRequestOwnerEmail: this.props.webpartContext.pageContext.user.loginName, tileRequestRequestedBy: this.props.webpartContext.pageContext.user.loginName });

            //Get TNumber from properties
            let tNumber: string = await this.pnpHelper.userProps("TNumber");

            //Get Id from lookup column
            sp.web.lists.getByTitle("UserMaster").items.filter(`Title eq '${tNumber}'`).select("Id").get().then(r => {
                this.setState({
                    tNumberId: (r[0].Id !== null || undefined) ? r[0].Id : 0
                });
                //return r[0].Id;
            });

            this.setState({ tNumber: tNumber });
            //console.log(`${tNumber} : ${this.state.tNumberId}`);
            this._refreshValidationData();
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "TileRequest", error, "_openTileRequestPanel");
        }
    }

    //call this function when tile request approved or rejected
    //this function will refresh the datasets for validation
    private _refreshValidationData = async () => {
        try {
            //populate application master data
            sp.web.lists.getByTitle(this.ApplicationMasterListName).items
                .select('Id', 'Link', 'IsActive')
                .getAll()
                .then((applicationTiles: IApplicationList[]) => {
                    this.allApplicationTiles = applicationTiles;
                });

            //get Tile request data for validaion
            sp.web.lists.getByTitle(this.lstTileRequest)
                .items.select('Id', 'Link')
                .filter(`ApprovalStatus eq 'Waiting for Approval'`)
                .top(5000)
                .get()
                .then((r): void => {
                    if (r.length > 0) {
                        this.allApplicationRequest = r;
                    }
                });
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "TileRequest", error, "_refreshValidationData");
        }
    }

    //to close panel
    private _hideTileRequestPanel = () => {
        try {
            this.setState({ tileRequestIsPanelOpen: false });

            //clear controls state
            this.clearTileReuqestControlsState();
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "TileRequest", error, "_hideTileRequestPanel");
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
                <div>
                    <PrimaryButton
                        onClick={this._saveTileRequest.bind(this)}
                        text={this.state.resourceListItems["add_tile_save_text"]}
                        title={this.state.resourceListItems["add_tile_save_text"]}
                    />{" "}
                    &nbsp;
                    <DefaultButton
                        onClick={this._hideTileRequestPanel.bind(this)}
                        text={this.state.resourceListItems["add_tile_cancel_text"]}
                        title={this.state.resourceListItems["add_tile_cancel_text"]}
                    />
                </div>
            </React.Fragment>
        );
    }

    // User clicks on save button, new item will create in "ApplicationRequests" list
    private _saveTileRequest = async () => {
        try {
            let isFormValid = this.validateTileReuqestControls();
            if (isFormValid) {
                this.setState({ tileRequestShowMessageBar: false });

                let itemTileRequest: ITileRequest = {
                    Title: this.state.tileRequestTitle,
                    Description: this.state.tileRequestDescription,
                    SearchKeywords: this.state.tileRequestKeywords,
                    Link: this.state.tileRequestUrlLink,
                    AvailableExternal: this.state.tileRequestAvailableExternal,
                    OwnerEmail: this.state.tileRequestOwnerEmail,
                    //RequestedById: this.state.tNumberId,
                    ColorCodeId: this.selectedColorId,
                    RequestedDate: new Date(),
                };

                //console.log(itemTileRequest);
                //Post data to ApplcationRequest List
                this.pnpHelper.createTileRequest(itemTileRequest);
                let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentActiveCollection["ID"]);
                this.props.callBackForLatestFollowedCollections(false, false, currentItemIndexInMyFollowedCollectionList, this.state.resourceListItems['add_tile_request_submitted_msg']);
                this._hideTileRequestPanel();
            } else {
                this.ctrTitle.current.focus();
                this.ctrDescription.current.focus();
                this.ctrKeywords.current.focus();
                this.ctrURL.current.focus();
                this.ctrTitle.current.focus();
                //show error message
                this.showError = this.selectedColorId === 0 ? true : false;
                //this.setState({ tileRequestShowMessageBar: true });
            }
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "TileRequest", error, "_saveTileRequest");
        }
    }

    //close MessageBar
    private _closeTileRequestMessageBar = () => {
        this.setState({ tileRequestShowMessageBar: false });
    }

    //validate controls state
    private validateTileReuqestControls = (): boolean => {
        let isValid = true;
        if (this.getTileRequestTitleErrorMessage(this.state.tileRequestTitle) !== "")
            isValid = false;
        if (this.getTileRequestDescriptionErrorMessage(this.state.tileRequestDescription) !== "")
            isValid = false;
        if (this.getTileRequestKeywordsErrorMessage(this.state.tileRequestKeywords) !== "")
            isValid = false;
        if (this.getTileRequestURLLinkErrorMessage(this.state.tileRequestUrlLink) !== "")
            isValid = false;
        if (this.selectedColorId === 0)
            isValid = false; this.showError = true;
        return isValid;
    }

    //clear controls state after panel close
    private clearTileReuqestControlsState = (): void => {
        this.setState({
            tileRequestTitle: "",
            tileRequestDescription: "",
            tileRequestKeywords: "",
            tileRequestUrlLink: "",
            tileRequestOwnerEmail: "",
            tileRequestIsUrlValid: false,
            tileRequestAvailableExternal: 1
        });
        this.selectedColorId = 0;
        this.showError = false;
    }
    //#endregion

    //#region Validation
    //Validate Title text
    private getTileRequestTitleErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value) ? this.state.resourceListItems['required_field_validation_message'] : "";// : value.length > 255 ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //Validate Description text
    private getTileRequestDescriptionErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value) ? this.state.resourceListItems['required_field_validation_message'] : value.length > 63999 ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }
    //Validate Keywords text
    private getTileRequestKeywordsErrorMessage = (value: string): string => {
        value = value.trim();
        return stringIsNullOrEmpty(value) ? this.state.resourceListItems['required_field_validation_message'] : value.length > 63999 ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
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
                        : this.allApplicationTiles.some(tile => {
                            if (tile.IsActive) {
                                return tile.Link.toLowerCase() === value.toLowerCase();
                            }
                        }) ? this.state.resourceListItems['validation_url_exist_text']
                            : this.allApplicationTiles.some(tile => {
                                if (!tile.IsActive) {
                                    return tile.Link.toLowerCase() === value.toLowerCase();
                                }
                            }) ? this.state.resourceListItems['validation_url_exist_disable_text']
                                : this.allApplicationRequest.some(tile => {
                                    return tile.Link.toLowerCase() === value.toLowerCase();
                                }) ? this.state.resourceListItems['add_tile_request_exist_msg'] : "";

        //to disable Test URL
        this.setState({
            tileRequestIsUrlValid: errMsg === "" ? true : false,
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
                : value.length > 63999 ? this.state.resourceListItems['validation_maximum_characters_text']
                    : emailPattern.test(value) ? ""
                        : this.state.resourceListItems['validation_invalid_email'];
        return errMsg;
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

    //capture color code
    private _onColorCodeChange = (
        event: React.FormEvent<HTMLDivElement>,
        item: IDropdownOption
    ) => {
        this.selectedColorId = parseInt(item.key.toString());
        //console.log(this.selectedColorId);
        this.setState({
            tileRequestColorCode: this.selectedColorId,
        });

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

    //capture OwnerEmail text
    private _onTileRequestOwnerEmailChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ) => {
        this.setState({ tileRequestOwnerEmail: newValue });
    }
    //#endregion

    //#endregion TileRequest Events and Validation Ends

    //#region Manage Breaking News Events and Validation

    //#region manage breaking news other events
    //Called when user clicked on "Add a tile" option from left pane
    private _openBreakingNewsPanel = () => {

        this.setState({ bNewsIsPanelOpen: true });

    }
    private _getBreakingNewsDetails = async () => {
        try {
            await sp.web.lists.getByTitle("BreakingNews").items.orderBy("Id", false).top(1).get().then(r => {
                if (r.length > 0) {
                    this.valDescription = r[0].Description;
                    this.varExpiryDate = new Date(r[0].ExpiryDate);
                    this.setState({
                        bNewsId: r[0].Id,
                        bNewsTitle: r[0].Title,
                        bNewsTitleFontColor: r[0].TitleFontColor,
                        bNewsDescription: r[0].Description,
                        bNewsBackgroundColor: r[0].BgColor,
                        bNewsContentFontColor: r[0].ContentFontColor,
                        bNewsExpiryDate: new Date(r[0].ExpiryDate),
                        bNewsIsActive: r[0].IsActive,
                        bNewsContentScrollSpeed: r[0].ContentScrollSpeed
                    });
                    //console.log(r[0]);
                }

            });
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "BreakingNews", error, "_getBreakingNewsDetails");
        }
    }
    //to close panel
    private _hideBreakingNewsPanel = () => {
        this.setState({ bNewsIsPanelOpen: false });

        //clear controls state
        //this.clearBreakingNewsControlsState();
    }
    //Populate panel footer content with "Save" and "Cancel" buttons
    private _onRenderBreakingNewsFooterContent = (): JSX.Element => {
        return (
            <React.Fragment>
                {this.state.bNewsShowMessageBar && (
                    <div>

                        <MessageBar
                            messageBarType={MessageBarType.error}
                            isMultiline={true}
                            onDismiss={this._closeBreakingNewsMessageBar.bind(this)}
                            dismissButtonAriaLabel="Close"
                        >
                            {this.state.saveControlsErrorMessage}
                        </MessageBar>
                        <br />
                    </div>
                )}
                <div>
                    <PrimaryButton
                        onClick={this._saveBreakingNews.bind(this)}
                        text={this.state.resourceListItems["bnews_save_text"]}
                    />{" "}
                    &nbsp;
                    <DefaultButton
                        onClick={this._hideBreakingNewsPanel.bind(this)}
                        text={this.state.resourceListItems["bnews_cancel_text"]}
                    />
                </div>
            </React.Fragment>
        );
    }
    // User clicks on save button, new item will create in "ApplicationRequests" list
    private _saveBreakingNews = () => {
        try {
            //alert(this.valDescription);
            this.setState({ bNewsDescription: this.valDescription });

            let isFormValid = this.validateBreakingNewsControls();
            if (isFormValid) {
                this.setState({ bNewsShowMessageBar: false });
                let itemBreakingNews: IBreakingNews = {
                    Title: this.state.bNewsTitle,
                    TitleFontColor: this.state.bNewsTitleFontColor,
                    Description: this.valDescription,
                    BgColor: this.state.bNewsBackgroundColor,
                    ContentFontColor: this.state.bNewsContentFontColor,
                    ContentScrollSpeed: this.state.bNewsContentScrollSpeed,
                    ExpiryDate: this.state.bNewsExpiryDate,
                    IsActive: this.state.bNewsIsActive,
                };

                if (this.state.bNewsId !== 0) {
                    //Update item to SPO list
                    sp.web.lists
                        .getByTitle("BreakingNews")
                        .items.getById(this.state.bNewsId).update(itemBreakingNews)
                        .then(() => {
                            this._hideBreakingNewsPanel();
                        });
                } else {
                    //Add new item to SPO list
                    sp.web.lists
                        .getByTitle("BreakingNews")
                        .items.add(itemBreakingNews)
                        .then(() => {
                            this._hideBreakingNewsPanel();
                        });
                }
                let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentActiveCollection["ID"]);
                this.props.callBackForLatestFollowedCollections(false, false, currentItemIndexInMyFollowedCollectionList, this.state.resourceListItems['bnews_request_submitted_msg']);

            } else {

                //this.ctrNewsBackgroundColor.current.focus();
                //this.ctrNewsContentFontColor.current.focus();

                this.setState({ bNewsShowMessageBar: true });
            }
        } catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "BreakingNews", error, "_saveBreakingNews");
        }
    }

    private validateBreakingNewsControls = (): boolean => {
        let errMsg = "Please validate: ";
        let isValid = true;
        if (this.getBreakingNewsTitleErrorMessage(this.state.bNewsTitle) !== "") {
            isValid = false;
            errMsg += "'Title' ";
        }
        if (this.getBreakingNewsTitleFontColorErrorMessage(this.state.bNewsTitleFontColor) !== "") {
            isValid = false;
            errMsg += "'Title Font Color' ";
        }
        if (this.valDescription === "<p><br></p>") {
            isValid = false;
            errMsg += "'Breaking News Description' ";
        }
        if (this.getBreakingNewsBgColorErrorMessage(this.state.bNewsBackgroundColor) !== "") {
            isValid = false;
            errMsg += "'Background Color' ";
        }
        if (this.getBreakingNewsFontColorErrorMessage(this.state.bNewsContentFontColor) !== "") {
            isValid = false;
            errMsg += "'Content Font Color' ";
        }
        if (this.getBreakingNewsExpiryDateErrorMessage(this.varExpiryDate) !== "") {
            isValid = false;
            errMsg += "'Expiry Date' ";
        }
        this.setState({ saveControlsErrorMessage: errMsg });
        return isValid;
    }

    //close MessageBar
    private _closeBreakingNewsMessageBar = () => {
        this.setState({ bNewsShowMessageBar: false });
    }
    //#endregion

    //#region validations
    private getBreakingNewsTitleErrorMessage = (value: string): string => {
        value = value.trim();
        return value == "" ? this.state.resourceListItems['required_field_validation_message'] : value.length > 255
            ? this.state.resourceListItems['validation_maximum_characters_text'] : "";
    }

    private getBreakingNewsTitleFontColorErrorMessage = (value: string): string => {
        value = value.trim();
        let regex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
        return value == "" ? this.state.resourceListItems['required_field_validation_message']
            : regex.test(value) ? "" : this.state.resourceListItems['validation_invalid_color'];
    }
    private getBreakingNewsBgColorErrorMessage = (value: string): string => {
        value = value.trim();
        let regex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
        return value == "" ? this.state.resourceListItems['required_field_validation_message']
            : regex.test(value) ? "" : this.state.resourceListItems['validation_invalid_color'];
    }

    private getBreakingNewsFontColorErrorMessage = (value: string): string => {
        value = value.trim();
        let regex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/;
        return value == "" ? this.state.resourceListItems['required_field_validation_message']
            : regex.test(value) ? "" : this.state.resourceListItems['validation_invalid_color'];
    }

    private getBreakingNewsExpiryDateErrorMessage = (value: Date): string => {
        this.varExpiryDate = value;
        this.setState({ bNewsExpiryDate: value });
        let msg = value.toString() == "" ? this.state.resourceListItems['required_field_validation_message']
            : value > new Date() ? ""
                : this.state.resourceListItems['validate_future_date'];

        //this.setState({ bNewsIsValidExpiryDate: (msg == "" ? true : false) });
        return msg;
    }

    //#endregion

    //#region text capture events
    /*onChange Events - Start */
    private _onBreakingNewsTitleChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newText: string
    ) => {
        this.setState({ bNewsTitle: newText });
    }

    private _onBreakingNewsTitleFontColorChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newText: string
    ) => {
        this.setState({ bNewsTitleFontColor: newText });
    }

    private _onBreakingNewsDescriptionChange = (newText: string) => {
        this.valDescription = newText;
        //this.setState({ bNewsDescription: newText });
        //console.log(this.valDescription);
        return newText;
    }

    private _onBreakingNewsBgColorChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newText: string
    ) => {
        this.setState({ bNewsBackgroundColor: newText });
    }

    private _onBreakingNewsFontColorChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newText: string
    ) => {
        this.setState({ bNewsContentFontColor: newText });
    }

    private _onBreakingNewsExpiryDateChange = (newText: Date) => {
        this.setState({ bNewsExpiryDate: newText });
        this.varExpiryDate = newText;
        //this.getBreakingNewsExpiryDateErrorMessage(newText)
    }

    private _onBreakingNewsIsActiveChange = (
        event: React.MouseEvent<HTMLElement>,
        newText: boolean
    ) => {
        this.setState({ bNewsIsActive: newText });
    }

    /*onChange Events - End */
    //#endregion

    //#endregion TileRequest Events and Validation Ends

}
