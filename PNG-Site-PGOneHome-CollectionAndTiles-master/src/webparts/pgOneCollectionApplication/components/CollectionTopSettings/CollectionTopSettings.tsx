import * as React from 'react';
import { ICollectionList } from "../Common/ICollectionList";
import { ICollectionRequest } from "../Common/ICollectionRequest";
import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Link } from 'office-ui-fabric-react/lib/Link';
import styles from './CollectionTopSettings.module.scss';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType, IPanelProps } from 'office-ui-fabric-react/lib/Panel';
import { IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { IApplicationList } from "../Common/IApplicationList";
import { PnPHelper } from '../PnPHelper/PnPHelper';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Text, ITextProps } from 'office-ui-fabric-react/lib/Text';
import { gAnalytics } from '../GA/gAnalytics';

export interface ICollectionTopSettingsProps {
    currentCollectionItem: ICollectionList;
    myFollowedCollectionsList: ICollectionList[];
    webpartContext: WebPartContext;
    matchedApplicationListBasedOnCollection: IApplicationList[];
    isCurrentCollectionBeingFollowed: boolean;
    resourceListItems: any[];
    callBackForLatestFollowedCollections: any;
    isOwnerForCurrentCollection: boolean;
    gAnalytics?: gAnalytics;
    loadReadonlyUserProfile: boolean;
}

export interface ICollectionTopSettingsState {
    currentCollectionItem: ICollectionList;
    myFollowedCollectionItems: ICollectionList[];
    currentHoverApplicationTile?: IApplicationList;
    isOwnerForCurrentCollection?: boolean;
    followOrUnfollowIconProps: IIconProps;
    followOrUnfollowText: string;
    hideUnfollowDialog: boolean;
    showEditSitesinCollectionPanel: boolean;
    showCollectionSettingsPanel: boolean;
    showEditSitesInstructionInPanel: boolean;
    applicationTiles: IApplicationList[];
    matchedApplicationListBasedOnCollection: IApplicationList[];
    isCurrentCollectionBeingFollowed: boolean;
    resourceListItems: any[];
    isCollectionChangeSettingsAlreadyRequested: boolean;
    collectionSettingsPanelLoading: boolean;
    hideDeleteCollectionConfirmationDialog: boolean;
    showCalloutMessage: boolean;
    editSitesInCollectionPanelLoading: boolean;
    valueRequiredErrorMessage: string;
    showRequiredErrorMessageForApplicationCheckBox: boolean;
    loadReadonlyUserProfile: boolean;
}

export class CollectionTopSettings extends React.Component<ICollectionTopSettingsProps, ICollectionTopSettingsState> {

    private pnpHelper: PnPHelper;
    private _allApplicationTiles: IApplicationList[] = [];
    private selectedApplicationInEditTilesCollectionPanel: IApplicationList[] = [];
    private unSelectedApplicationInEditTilesCollectionPanel: IApplicationList[] = [];
    private errTitle: string = "PNG-Site-PGOneHome-CollectionAndTiles";
    private errModule: string = "CollectionsLeftNavigation.tsx";

    constructor(props: ICollectionTopSettingsProps, state: ICollectionTopSettingsState) {
        super(props);
        this.state = {
            currentCollectionItem: this.props.currentCollectionItem,
            myFollowedCollectionItems: this.props.myFollowedCollectionsList,
            isOwnerForCurrentCollection: this.props.isOwnerForCurrentCollection,
            followOrUnfollowIconProps: { iconName: 'CheckMark' },
            followOrUnfollowText: this.props.resourceListItems["following_this_collection"],
            hideUnfollowDialog: true,
            showCollectionSettingsPanel: false,
            showEditSitesinCollectionPanel: false,
            applicationTiles: [],
            showEditSitesInstructionInPanel: true,
            currentHoverApplicationTile: { Title: '', ID: '', Link: '', Description: '', OwnerEmail: '', SearchKeywords: '', ColorCode: { BgColor: '', ForeColor: '' } },
            matchedApplicationListBasedOnCollection: this.props.matchedApplicationListBasedOnCollection,
            isCurrentCollectionBeingFollowed: this.props.isCurrentCollectionBeingFollowed,
            resourceListItems: this.props.resourceListItems,
            isCollectionChangeSettingsAlreadyRequested: true,
            hideDeleteCollectionConfirmationDialog: true,
            collectionSettingsPanelLoading: true,
            showCalloutMessage: false,
            editSitesInCollectionPanelLoading: true,
            valueRequiredErrorMessage: "",
            showRequiredErrorMessageForApplicationCheckBox: false,
            loadReadonlyUserProfile: this.props.loadReadonlyUserProfile,
        };
        this.pnpHelper = new PnPHelper(this.props.webpartContext);
    }
    //Component will receive properties from Parent
    public componentWillReceiveProps(newProps: ICollectionTopSettingsProps) {
        this.setState({
            currentCollectionItem: newProps.currentCollectionItem,
            myFollowedCollectionItems: newProps.myFollowedCollectionsList,
            isOwnerForCurrentCollection: newProps.isOwnerForCurrentCollection,
            matchedApplicationListBasedOnCollection: newProps.matchedApplicationListBasedOnCollection,
            isCurrentCollectionBeingFollowed: newProps.isCurrentCollectionBeingFollowed,
            resourceListItems: newProps.resourceListItems,
            followOrUnfollowText: newProps.resourceListItems["following_this_collection"],
            loadReadonlyUserProfile: newProps.loadReadonlyUserProfile,
        });
    }

    public render() {
        return (
            <div className={`ms-Grid-row ${styles.collectionTopSettings}`}>
                <div className={styles.collectionDetails}>
                    <h2 className={`${styles.collectionName}`}>
                        {this.props.currentCollectionItem.Title}
                    </h2>
                    {!this.state.loadReadonlyUserProfile &&
                        <div className={styles.followorUnfollowSection}>
                            {(this.state.isCurrentCollectionBeingFollowed)
                                ? //show following button if it is already being followed
                                <div className={styles.followingorUnfollow}>
                                    {(!this.state.isOwnerForCurrentCollection)
                                        &&
                                        <div>
                                            <DefaultButton
                                                className={`ms-hiddenMdDown ${styles.followingOrUnfollowButton} ${this.state.followOrUnfollowText == this.state.resourceListItems["following_this_collection"] ? styles.followingButton : styles.unfollowButton}`}
                                                text={this.state.followOrUnfollowText}
                                                iconProps={this.state.followOrUnfollowIconProps}
                                                onClick={this._showUnfollowDialog.bind(this)}
                                                onMouseEnter={this._toggleFollowTextHover.bind(this)}
                                                onMouseLeave={this._toggleUnFollowTextHover.bind(this)}
                                                title={this.state.resourceListItems["unfollow_collection_text"]}
                                            />
                                            <IconButton
                                                className={`ms-hiddenLgUp ${styles.followingOrUnfollowButton} ${this.state.followOrUnfollowText == this.state.resourceListItems["following_this_collection"] ? styles.followingButton : styles.unfollowButton}`}
                                                iconProps={this.state.followOrUnfollowIconProps}
                                                onClick={this._showUnfollowDialog.bind(this)}
                                                onMouseEnter={this._toggleFollowTextHover.bind(this)}
                                                onMouseLeave={this._toggleUnFollowTextHover.bind(this)}
                                                onTouchStart={this._toggleFollowTextHover.bind(this)}
                                                onTouchEnd={this._toggleUnFollowTextHover.bind(this)}
                                            />
                                        </div>
                                    }
                                </div>
                                : //show follow button if it is not currently followed
                                <div className={styles.follow}>
                                    <DefaultButton
                                        className={`ms-hiddenMdDown ${styles.followButton}`}
                                        text={this.state.resourceListItems["follow_this_collection"]}
                                        title={this.state.resourceListItems["follow_this_collection"]}
                                        iconProps={{ iconName: 'Add' }}
                                        onClick={this._followCollection.bind(this)}
                                    />
                                    <IconButton
                                        className={`ms-hiddenLgUp ${styles.followButton}`}
                                        iconProps={{ iconName: 'Add' }}
                                        onClick={this._followCollection.bind(this)}
                                    />
                                </div>
                            }
                        </div>
                    }
                </div>

                <div className={styles.collectionSettings}>
                    <div className={styles.shareCollectionInMail}>
                        <a href="#"
                            onClick={() => {
                                this.props.gAnalytics.shareCollectionClick(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);
                                location.href = "mailto:?subject=Check out this collection&body=" + this.props.webpartContext.pageContext.site.absoluteUrl + "?collectionId=" + this.state.currentCollectionItem.ID;
                            }}
                            title={this.state.resourceListItems["share_collection"]}
                        >
                            <Icon iconName="Mail" />
                            <span className={"ms-hiddenMdDown"}>{this.state.resourceListItems["share_collection"]}</span>
                        </a>
                    </div>
                    {this.state.isOwnerForCurrentCollection &&
                        //show owner related settings only if he is owner of the current collection
                        <div className={styles.ownerTasks}>
                            <span className={"ms-hiddenMdDown"}>{this.state.resourceListItems["owner_tasks"]}</span>
                            <Link onClick={this._openEditSitesinCollectionPanel.bind(this)} title={this.state.resourceListItems["edit_sites_in_collection"]}>
                                <Icon iconName="Edit" />
                                <span className={"ms-hiddenMdDown"}>{this.state.resourceListItems["edit_sites_in_collection"]}</span>
                            </Link>
                            <Link onClick={this._showCollectionSettingsDialog.bind(this)} title={this.state.resourceListItems["edit_collection_properties"]}>
                                <Icon iconName="Settings" />
                                <span className={"ms-hiddenMdDown"}>{this.state.resourceListItems["edit_collection_properties"]}</span>
                            </Link>
                        </div>
                    }
                </div>
                {/* UnFollow Dialog Confirmation*/}
                <Dialog
                    className={styles.followDialog}
                    hidden={this.state.hideUnfollowDialog}
                    onDismiss={this._hideDialog.bind(this)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        showCloseButton: false
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 650 } }
                    }}
                >
                    <div className={styles.followDialogContent}>
                        <Icon iconName="Info" />
                        <p className={styles.followDialogTitle}>{this.state.resourceListItems["unfollow_Collection_Dialog_Title"]}</p>
                        <p className={styles.followDialogText}>{this.state.resourceListItems["unfollow_Collection_Dialog_Text"]}</p>
                    </div>
                    <DialogFooter>
                        <DefaultButton onClick={this._hideDialog.bind(this)} title={this.state.resourceListItems["cancel"]} text={this.state.resourceListItems["cancel"]} />
                        <PrimaryButton className={styles.unFollowbutton} onClick={this._unFollowCollection.bind(this)} title={this.state.resourceListItems["unfollow_Collection_Dialog_Unfollow_Button_Text"]} text={this.state.resourceListItems["unfollow_Collection_Dialog_Unfollow_Button_Text"]} />
                    </DialogFooter>
                </Dialog>

                {/* Edit Sites in this collection */}
                <Panel
                    className={styles.editSitesinCollectionPanel}
                    isOpen={this.state.showEditSitesinCollectionPanel}
                    onDismiss={this._hidePanel}
                    type={PanelType.large}
                    closeButtonAriaLabel={this.state.resourceListItems["close"]}
                    headerText={this.state.resourceListItems["edit_sites_in_collection"]}
                    onRenderFooterContent={this._onRenderFooterContentEditSitesinCollection}
                    headerClassName={styles.headerText}
                >
                    <div className={`ms-hiddenLgUp ${styles.infoIconArea}`}>
                        <Icon iconName="Info"
                            onClick={() => { this.setState({ showCalloutMessage: !this.state.showCalloutMessage }); }}
                        />
                    </div>
                    {this.state.showCalloutMessage && //only call out display for mobiles
                        <Callout
                            role="alertdialog"
                            gapSpace={0}
                            setInitialFocus
                            target={`.${styles.infoIconArea}`}
                            className={styles.callout}
                            onDismiss={() => { this.setState({ showCalloutMessage: !this.state.showCalloutMessage }); }}
                        >
                            <div className={styles.editSitesInCollectionCallOutContent}>
                                <p>{this.state.resourceListItems["edit_sites_in_collection_instruction_line1_for_mobile"]}</p>
                                <p>{this.state.resourceListItems["edit_sites_in_collection_instruction_line2_for_mobile"]}</p>
                            </div>
                        </Callout>
                    }

                    <div className={`ms-Grid ${styles.editSitesinCollectionContent}`}>
                        {this.state.editSitesInCollectionPanelLoading
                            ? //show loading until API call success
                            <div className={`ms-Grid-row`}>
                                <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.spinner}`}>
                                    <Spinner size={SpinnerSize.medium} label="Loading..." ariaLive="assertive" labelPosition="right" />
                                </div>
                            </div>
                            :
                            <div className={`ms-Grid-row`}>
                                <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg4 ${styles.editSitesinCollectionLeftPart}`}>
                                    <SearchBox
                                        placeholder="Search Application Tiles"
                                        onChange={this._searchApplicationList.bind(this)}
                                        onSearch={this._searchApplicationList.bind(this)}
                                    />
                                    {this.state.showRequiredErrorMessageForApplicationCheckBox && <Label className={styles.requiredCheckBoxCheck}>{this.state.resourceListItems["edit_sites_in_collection_error_message"]}</Label>}
                                    <ul className={styles.applicationTileItems} aria-label={this.state.resourceListItems["create_collection_select_tile_label"]}>
                                        {
                                            this.state.applicationTiles.map((applicationTile: any, i: any) =>
                                                <li title={applicationTile.Title} onMouseOver={this._hoverApplicationList.bind(this, applicationTile)} onFocus={this._hoverApplicationList.bind(this, applicationTile)}>
                                                    <Checkbox
                                                        title={applicationTile.Title}
                                                        label={applicationTile.Title}
                                                        checked={
                                                            this.state.matchedApplicationListBasedOnCollection.some((matchedApplication: any) => {
                                                                return matchedApplication.ID === applicationTile.ID;
                                                            })
                                                        }
                                                        onChange={(event) => this._handleApplicationTileSelection(event, applicationTile)}
                                                    />
                                                </li>
                                            )
                                        }
                                    </ul>
                                    {this._allApplicationTiles.length != this.state.applicationTiles.length
                                            &&
                                            <div aria-live="polite" className="sr-only" role="status">{this.state.applicationTiles.length} suggestions found</div>
                                    }
                                </div>
                                <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg8 ${styles.editSitesinCollectionRightPart}`}>
                                    <div className={this.state.showEditSitesInstructionInPanel ? styles.showEditSitesInCollectionInstructionInPanel : styles.hideEditSitesInCollectionInstructionInPanel}>
                                        <p>{this.state.resourceListItems["edit_sites_in_collection_instruction_line1_for_desktop"]}</p>
                                        <p>{this.state.resourceListItems["edit_sites_in_collection_instruction_line2_for_desktop"]}</p>
                                    </div>
                                    <div className={this.state.showEditSitesInstructionInPanel ? styles.hideEditSitesInCollectionInstructionInPanel : styles.showEditSitesInCollectionInstructionInPanel}>
                                        <div className={styles.tileColor} style={{ backgroundColor: this.state.currentHoverApplicationTile.ColorCode.BgColor }} />
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
                        }
                    </div>
                </Panel>

                {/* Collection Settings Panel */}
                <Dialog
                    className={styles.deleteCollectionConfirmationDialog}
                    hidden={this.state.hideDeleteCollectionConfirmationDialog}
                    onDismiss={this._hideDialog.bind(this)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        showCloseButton: false
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 650 } }
                    }}
                >
                    <div className={styles.deleteCollectionConfirmationDialogContent}>
                        <Icon iconName="Info" />
                        <p className={styles.deleteCollectionConfirmationDialogTitle}>{this.state.resourceListItems["delete_Collection_Dialog_Title"]}</p>
                        <p className={styles.deleteCollectionConfirmationDialogText}>{this.state.resourceListItems["delete_Collection_Dialog_Text"]}</p>
                    </div>
                    <DialogFooter>
                        <DefaultButton onClick={this._hideDialog.bind(this)} title={this.state.resourceListItems["cancel"]} text={this.state.resourceListItems["cancel"]} />
                        <PrimaryButton className={styles.deleteCollectionbutton} onClick={this._deleteCollectionAndItsApplications.bind(this)} title={this.state.resourceListItems["delete_Collection_Dialog_delete_Button_Text"]} text={this.state.resourceListItems["delete_Collection_Dialog_delete_Button_Text"]} />
                    </DialogFooter>
                </Dialog>
                <Panel
                    className={styles.collectionSettingsPanel}
                    isOpen={this.state.showCollectionSettingsPanel}
                    onDismiss={this._hidePanel}
                    type={PanelType.medium}
                    closeButtonAriaLabel={this.state.resourceListItems["close"]}
                    headerText={this.state.resourceListItems["edit_collection_header"]}
                    onRenderFooterContent={this._onRenderFooterContentCollectionSettings}
                    headerClassName={styles.headerText}
                >
                    <div className={styles.collectionSettingsContent}>
                        {this.state.collectionSettingsPanelLoading
                            ? //show spinner until API success
                            <div className={styles.spinner}>
                                <Spinner size={SpinnerSize.medium} label="Loading..." ariaLive="assertive" labelPosition="right" />
                            </div>
                            : <div>
                                <TextField required label={this.state.resourceListItems["new_collection_name"]} value={this.state.currentCollectionItem.Title} onChange={this._collectionNameChange.bind(this)} errorMessage={this.state.currentCollectionItem.Title.trim() === "" ? this.state.valueRequiredErrorMessage : ""} maxLength={255} disabled={this.state.isCollectionChangeSettingsAlreadyRequested} />
                                {/* <Checkbox label={this.state.resourceListItems["new_collection_make_public"]} onChange={this.isPublicCollectionCheckBoxChange.bind(this)} checked={this.state.currentCollectionItem.PublicCollection == 0 ? false : true} disabled={this.state.isCollectionChangeSettingsAlreadyRequested || this.state.currentCollectionItem.CorporateCollection == 1} /> */}
                                {this.state.currentCollectionItem.CorporateCollection == 1 && <Text className={styles.corpCollectionMessage}>{this.state.resourceListItems["corporateCollectionMesageInEditCollectionSettings"]}</Text>}
                                <TextField required label={this.state.resourceListItems["new_collection_description"]} onChange={this._collectionDescriptionChange.bind(this)} value={this.state.currentCollectionItem.Description} multiline rows={3} errorMessage={this.state.currentCollectionItem.Description.trim() === "" ? this.state.valueRequiredErrorMessage : ""} maxLength={2000} disabled={this.state.isCollectionChangeSettingsAlreadyRequested} />
                                {this.state.isCollectionChangeSettingsAlreadyRequested &&
                                    <span className={styles.adminRequestAvailableMessage}>{this.state.resourceListItems["collection_settings_awaiting_admin_approval_message"]}</span>}
                            </div>
                        }
                    </div>
                </Panel>
            </div >
        );
    }

    private _onRenderFooterContentCollectionSettings = (props: IPanelProps, defaultRender: IRenderFunction<IPanelProps>): JSX.Element => {
        return (
            <React.Fragment>
                <div className={styles.btnContainer}>
                    <PrimaryButton onClick={this._saveCollectionSettings.bind(this)} title={this.state.resourceListItems["save"]} text={this.state.resourceListItems["save"]} disabled={this.state.isCollectionChangeSettingsAlreadyRequested} />
                    <DefaultButton onClick={this._hidePanel.bind(this)} title={this.state.resourceListItems["cancel"]} text={this.state.resourceListItems["cancel"]} disabled={this.state.isCollectionChangeSettingsAlreadyRequested} />
                    {this.state.currentCollectionItem.UnDeletable == 0 && //show delete button only Undeleteable flag is 0
                        <DefaultButton onClick={() => { this.setState({ hideDeleteCollectionConfirmationDialog: false }); }} title={this.state.resourceListItems["delete"]} text={this.state.resourceListItems["delete"]} disabled={this.state.isCollectionChangeSettingsAlreadyRequested} />}
                </div>
            </React.Fragment >
        );
    }

    private _onRenderFooterContentEditSitesinCollection = (props: IPanelProps, defaultRender: IRenderFunction<IPanelProps>): JSX.Element => {
        return (
            <React.Fragment>
                <div className={styles.btnContainer}>
                    <PrimaryButton onClick={this._saveSitesInCollection.bind(this)} title={this.state.resourceListItems["save"]} text={this.state.resourceListItems["save"]} />
                    <DefaultButton onClick={this._hidePanel.bind(this)} title={this.state.resourceListItems["cancel"]} text={this.state.resourceListItems["cancel"]} />
                </div>
            </React.Fragment >
        );
    }

    // show current application details on Hover of Application list
    private _hoverApplicationList = (currentHoverApplication: IApplicationList): void => {
        this.setState({
            currentHoverApplicationTile: currentHoverApplication,
            showEditSitesInstructionInPanel: false
        });

    }

    private _toggleFollowTextHover = (): void => {
        this.setState({
            followOrUnfollowText: this.state.resourceListItems["unfollow_collection_text"],
            followOrUnfollowIconProps: { iconName: 'Cancel' }
        });
    }

    private _toggleUnFollowTextHover = (): void => {
        this.setState({
            followOrUnfollowText: this.state.resourceListItems["following_this_collection"],
            followOrUnfollowIconProps: { iconName: 'CheckMark' }
        });
    }

    private _showUnfollowDialog = (): void => {
        this.setState({ hideUnfollowDialog: false });
    }

    //To Show Collection Settings current value in the panel
    private _showCollectionSettingsDialog = (): void => {
        try {
            //Google Analytics: collection Settings Called
            this.props.gAnalytics.collectionSettingsCalled(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);

            this.setState({ showCollectionSettingsPanel: true, collectionSettingsPanelLoading: true });
            this.pnpHelper.getPublicCollectionRequestDetailsBasedOnCollectionID(this.state.currentCollectionItem)
                .then((collectionRequestAvailable) => {
                    this.setState({
                        isCollectionChangeSettingsAlreadyRequested: collectionRequestAvailable,
                        collectionSettingsPanelLoading: false
                    });
                });
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _showCollectionSettingsDialog()", error, "Open Collection Settings Panel");
        }
    }

    private _openEditSitesinCollectionPanel = (): void => {
        try {
            //Google Analytics: edit Tiles In Collection Clicked
            this.props.gAnalytics.editSitesInCollectionClicked(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);
            this.setState({ showEditSitesinCollectionPanel: true });

            this.setState({
                showEditSitesinCollectionPanel: true,
                editSitesInCollectionPanelLoading: true
            });
            this.pnpHelper.getApplicationTiles()
                .then((applicationTiles: IApplicationList[]) => {
                    this._allApplicationTiles = applicationTiles;
                    this.setState({
                        applicationTiles: applicationTiles,
                        editSitesInCollectionPanelLoading: false,
                    });
                });
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _openEditSitesinCollectionPanel()", error, "Open Edit Sites in Collection Panel");
        }
    }

    private _handleApplicationTileSelection(event: any,
        applicationTileCheckBoxValue: any) {
        try {
            if (event.target.checked) {
                //Google Analytics: tile checked from list (Create a new collection)
                this.props.gAnalytics.appListAdd(
                    applicationTileCheckBoxValue.Title,
                    applicationTileCheckBoxValue.Link,
                    this.state.currentCollectionItem.Title,
                    this.state.currentCollectionItem.ID,
                    applicationTileCheckBoxValue.ID);
                this.selectedApplicationInEditTilesCollectionPanel.push(applicationTileCheckBoxValue);
                var pushedCurrentApplicationToSelectedList = this.state.matchedApplicationListBasedOnCollection.concat(this.selectedApplicationInEditTilesCollectionPanel);
                this.setState({
                    // just to remove duplicates before setting the state
                    matchedApplicationListBasedOnCollection: pushedCurrentApplicationToSelectedList.filter((item: any, pos: any, self: any) => {
                        return self.indexOf(item) == pos;
                    }),
                    showRequiredErrorMessageForApplicationCheckBox: false,
                });
                // push only newly added records
                this.selectedApplicationInEditTilesCollectionPanel = this.selectedApplicationInEditTilesCollectionPanel.filter((selectedApplicationItem: any) => {
                    return !this.props.matchedApplicationListBasedOnCollection.some((alreadyAddedApplicationItem: any) => {
                        return alreadyAddedApplicationItem.ID === selectedApplicationItem.ID;
                    });
                });

                this.unSelectedApplicationInEditTilesCollectionPanel = this.unSelectedApplicationInEditTilesCollectionPanel
                    .filter((unSelectedApplicationItem: any) => {
                        return unSelectedApplicationItem.ID !== applicationTileCheckBoxValue.ID;
                    });
            }
            else if (!event.target.checked) {
                //Google Analytics: tile unchecked from list (Create a new collection)
                this.props.gAnalytics.appListRemove(
                    applicationTileCheckBoxValue.Title,
                    applicationTileCheckBoxValue.Link,
                    this.state.currentCollectionItem.Title,
                    this.state.currentCollectionItem.ID,
                    applicationTileCheckBoxValue.ID);

                this.selectedApplicationInEditTilesCollectionPanel = this.selectedApplicationInEditTilesCollectionPanel
                    .filter((unSelectedApplicationItem: any) => {
                        return unSelectedApplicationItem.ID !== applicationTileCheckBoxValue.ID;
                    });
                this.unSelectedApplicationInEditTilesCollectionPanel.push(applicationTileCheckBoxValue);
                //// delete only already followed records
                this.unSelectedApplicationInEditTilesCollectionPanel = this.unSelectedApplicationInEditTilesCollectionPanel.filter((unSelectedApplicationItem: any) => {
                    return this.props.matchedApplicationListBasedOnCollection.some((alreadyAddedApplicationItem: any) => {
                        return alreadyAddedApplicationItem.ID === unSelectedApplicationItem.ID;
                    });
                });

                this.unSelectedApplicationInEditTilesCollectionPanel = this.unSelectedApplicationInEditTilesCollectionPanel.map((value: any) => {
                    this.props.matchedApplicationListBasedOnCollection.map((alreadyAddedApplicationItem: any) => {
                        if (alreadyAddedApplicationItem.ID == value.ID) {
                            value["ApplicationCollectionMatrixID"] = alreadyAddedApplicationItem["ApplicationCollectionMatrixID"];
                        }
                    });
                    return value;
                });

                this.unSelectedApplicationInEditTilesCollectionPanel = this.unSelectedApplicationInEditTilesCollectionPanel.filter((item: any, pos: any, self: any) => {
                    return self.indexOf(item) == pos;
                });
                this.setState({
                    matchedApplicationListBasedOnCollection: this.state.matchedApplicationListBasedOnCollection.filter((unSelectedApplicationItem: any) => {
                        return unSelectedApplicationItem.ID !== applicationTileCheckBoxValue.ID;
                    })
                });
            }
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _handleApplicationTileSelection()", error, "Handle Application Tile Selection in Edit Tiles in Collection Panel");
        }
    }
    private _hideDialog = (): void => {
        this.setState({
            hideUnfollowDialog: true,
            hideDeleteCollectionConfirmationDialog: true
        });
    }

    private _hidePanel = (): void => {
        this.setState({
            showEditSitesinCollectionPanel: false,
            showCollectionSettingsPanel: false,
            showEditSitesInstructionInPanel: true,
            applicationTiles: this._allApplicationTiles,
            currentCollectionItem: this.props.currentCollectionItem,
            showRequiredErrorMessageForApplicationCheckBox: false,
            matchedApplicationListBasedOnCollection: this.props.matchedApplicationListBasedOnCollection
        });
    }

    private _searchApplicationList = (searchValue: string): void => {
        try {
            //Google Analytics: tile search from list (Create a new collection)
            this.props.gAnalytics.appListSearch(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID, searchValue);

            let filteredApplicationsList: IApplicationList[] = [];
            this._allApplicationTiles.map((data: any, i: any) => {
                if (String(data["Title"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(data["Description"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1 ||
                    String(data["SearchKeywords"]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1) {
                    filteredApplicationsList.push(data);
                }
            });
            this.setState({
                applicationTiles: filteredApplicationsList
            });
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _searchApplicationList()", error, "Search Application List in Edit Tiles in Collection Panel");
        }
    }

    private _collectionNameChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ): void => {
        let currentCollectionItem = { ...this.state.currentCollectionItem };
        currentCollectionItem["Title"] = newValue;
        this.setState({
            currentCollectionItem: currentCollectionItem
        });
    }

    private _collectionDescriptionChange = (
        event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
        newValue: string
    ): void => {
        let currentCollectionItem = { ...this.state.currentCollectionItem };
        currentCollectionItem["Description"] = newValue;
        this.setState({
            currentCollectionItem: currentCollectionItem
        });
    }

    private isPublicCollectionCheckBoxChange = (
        event: React.FormEvent<HTMLElement>,
        checked: boolean): void => {
        let currentCollectionItem = { ...this.state.currentCollectionItem };
        currentCollectionItem["PublicCollection"] = Number(checked);
        this.setState({
            currentCollectionItem: currentCollectionItem
        });
    }

    private _deleteCollectionAndItsApplications = async (): Promise<void> => {
        try {
            if (this.state.currentCollectionItem.PublicCollection == 0) {
                Promise.all([
                    // UnCommented the below first call due to Restrict Deletion betwen Collection and its Applications. Need to explicitly delete application from collection application matrix table.
                    // this.pnpHelper.deleteItemsOnCollectionApplicationMatrixList(this.props.matchedApplicationListBasedOnCollection),
                    this.pnpHelper.deleteCollectionItemOnCollectionMasterList(this.props.currentCollectionItem)
                ]).then(() => {
                    this.props.callBackForLatestFollowedCollections(false, true, 0);
                });
            }
            else {
                let currentUserTNumber = await this.pnpHelper.userProps("TNumber");
                let currentUserItem = await this.pnpHelper.getCurrentUserItemID(currentUserTNumber);
                let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentCollectionItem.ID);
                let collectionRequestItem: ICollectionRequest = {
                    Title: this.state.currentCollectionItem.Title,
                    Description: this.state.currentCollectionItem.Description,
                    PublicCollection: this.state.currentCollectionItem.PublicCollection,
                    DefaultMyCollection: this.state.currentCollectionItem.DefaultMyCollection,
                    CorporateCollection: this.state.currentCollectionItem.CorporateCollection,
                    StandardOrder: this.state.currentCollectionItem.StandardOrder,
                    UnDeletable: this.state.currentCollectionItem.UnDeletable,
                    ExistingItemID: Number(this.state.currentCollectionItem.ID),
                    CollectionOwnerId: currentUserItem[0].ID,
                    RequestedAction: "Deletion"
                };
                this.pnpHelper.addPublicCollectionsToRequestList(collectionRequestItem).then(() => {
                    this.props.callBackForLatestFollowedCollections(false, false, currentItemIndexInMyFollowedCollectionList, this.state.resourceListItems["collection_settings_public_collection_deletion_message"]);
                });
            }

            this.setState({ showCollectionSettingsPanel: false });
            this.props.callBackForLatestFollowedCollections(true);
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _deleteCollectionAndItsApplications()", error, "Delete a Collection");
        }
    }
    private _saveCollectionSettings = async (): Promise<void> => {
        try {
            //Google Analytics: collection Settings Save Called
            this.props.gAnalytics.collectionSettingsSave(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);

            if (this.state.currentCollectionItem.Title.trim() === "" || this.state.currentCollectionItem.Description.trim() === "") {
                this.setState({
                    showCollectionSettingsPanel: true,
                    valueRequiredErrorMessage: this.state.resourceListItems["required_field_validation_message"]
                });
            }
            else if (this.state.currentCollectionItem.Title.trim() == this.props.currentCollectionItem.Title.trim()
                && this.state.currentCollectionItem.Description.trim() == this.props.currentCollectionItem.Description.trim()
                && this.state.currentCollectionItem.PublicCollection == this.props.currentCollectionItem.PublicCollection) {
                this._hidePanel();
            }
            else {
                // Creating Collection Request for Private to Public and Vice Versa
                if ((this.props.currentCollectionItem.PublicCollection !== this.state.currentCollectionItem.PublicCollection) || this.state.currentCollectionItem.PublicCollection == 1) {
                    let currentUserTNumber = await this.pnpHelper.userProps("TNumber");
                    let currentUserItem = await this.pnpHelper.getCurrentUserItemID(currentUserTNumber);
                    let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentCollectionItem.ID);
                    let collectionRequestItem: ICollectionRequest = {
                        Title: this.state.currentCollectionItem.Title,
                        Description: this.state.currentCollectionItem.Description,
                        PublicCollection: this.state.currentCollectionItem.PublicCollection,
                        DefaultMyCollection: this.state.currentCollectionItem.DefaultMyCollection,
                        CorporateCollection: this.state.currentCollectionItem.CorporateCollection,
                        StandardOrder: this.state.currentCollectionItem.StandardOrder,
                        UnDeletable: this.state.currentCollectionItem.UnDeletable,
                        ExistingItemID: Number(this.state.currentCollectionItem.ID),
                        CollectionOwnerId: currentUserItem[0].ID,
                        RequestedAction: "Modification"
                    };
                    let successMessage = this.props.currentCollectionItem.PublicCollection == 0 ? this.state.resourceListItems["collection_settings_private_to_public_collection_message"] : this.state.resourceListItems["collection_settings_public_collection_change_message"];

                    this.pnpHelper.addPublicCollectionsToRequestList(collectionRequestItem).then(() => {
                        this.props.callBackForLatestFollowedCollections(false, false, currentItemIndexInMyFollowedCollectionList, successMessage);
                    });
                }
                // Directly update the record in the Collection Master
                else if (this.state.currentCollectionItem.PublicCollection == 0) {
                    let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentCollectionItem.ID);
                    let collectionItem: ICollectionList = {
                        ID: this.state.currentCollectionItem.ID,
                        Title: this.state.currentCollectionItem.Title,
                        Description: this.state.currentCollectionItem.Description,
                        PublicCollection: this.state.currentCollectionItem.PublicCollection,
                    };
                    this.pnpHelper.updateCollectionSettings(collectionItem).then(() => {
                        this.props.callBackForLatestFollowedCollections(false, true, currentItemIndexInMyFollowedCollectionList);
                    });
                }

                this.setState({ showCollectionSettingsPanel: false });
                this.props.callBackForLatestFollowedCollections(true);
            }
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _saveCollectionSettings()", error, "Save Collection Settings");
        }
    }

    private _saveSitesInCollection = async (): Promise<void> => {
        try {
            //Google Analytics: edit Tiles In Collection Save Clicked
            this.props.gAnalytics.editSitesInCollectionSaved(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);

            if (this.state.matchedApplicationListBasedOnCollection.length == 0) {
                this.setState({ showRequiredErrorMessageForApplicationCheckBox: true });
            }
            else if (this.selectedApplicationInEditTilesCollectionPanel.length == 0 && this.unSelectedApplicationInEditTilesCollectionPanel.length == 0) {
                this._hidePanel();
            }
            else {
                // take the currender maximun order and add this item at the last
                let currentMaximumAppOrder = Math.max.apply(Math, this.props.matchedApplicationListBasedOnCollection.map((applicationTileItem: any) => { return applicationTileItem.AppOrder; }));
                let currentItemIndexInMyFollowedCollectionList = this.state.myFollowedCollectionItems.map((x) => { return x.ID; }).indexOf(this.state.currentCollectionItem.ID);
                Promise.all([
                    //add selected items in the CollectionApplicationMatrix List
                    this.pnpHelper.addItemToCollectionApplicationMatrixList(this.selectedApplicationInEditTilesCollectionPanel, this.state.currentCollectionItem.ID, currentMaximumAppOrder),
                    //delete unselecteditems in the CollectionApplicationMatrix List
                    this.pnpHelper.deleteItemsOnCollectionApplicationMatrixList(this.unSelectedApplicationInEditTilesCollectionPanel)
                ]).then(() => {
                    this.props.callBackForLatestFollowedCollections(false, true, currentItemIndexInMyFollowedCollectionList);
                });
                this.setState({ showEditSitesinCollectionPanel: false });
                this.props.callBackForLatestFollowedCollections(true);
            }
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _saveSitesInCollection()", error, "Save Sites In Collection");
        }
    }

    private _followCollection = async (): Promise<void> => {
        try {
            //Google Analytics: Collection followed with Direct link 
            this.props.gAnalytics.collectionDirectFollow(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);

            let currentUserTNumber = await this.pnpHelper.userProps("TNumber");
            let currentUserItem = await this.pnpHelper.getCurrentUserItemID(currentUserTNumber);
            let currentMaximumCollectionOrder = Math.max.apply(Math, this.state.myFollowedCollectionItems.map((followedCollectionItem: any) => { return followedCollectionItem.CollectionOrder; }));
            let currentMyFollowedCollectionsLength = this.state.myFollowedCollectionItems.length;
            let currentCollectionItem = [];
            currentCollectionItem.push(this.state.currentCollectionItem);

            //add the current collection to the User Collection Matrix List
            this.pnpHelper.addMulipleItemsToUserCollectionMatrixList(currentUserItem, currentCollectionItem, currentMaximumCollectionOrder)
                .then(() => {
                    this.props.callBackForLatestFollowedCollections(false, true, currentMyFollowedCollectionsLength);
                });
            this.props.callBackForLatestFollowedCollections(true);
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _followCollection()", error, "Follow Collection");
        }
    }

    private _unFollowCollection = (): void => {
        try {
            //Google Analytics: Collection unfollowed with Direct link 
            this.props.gAnalytics.collectionDirectUnfollow(this.state.currentCollectionItem.Title, this.state.currentCollectionItem.ID);

            let currentCollectionItem = [];
            currentCollectionItem.push(this.state.currentCollectionItem);
            // remove the item from User Collection Matrix List
            this.pnpHelper.deleteItemsOnUserCollectionMatrixList(currentCollectionItem).then(() => {
                this.props.callBackForLatestFollowedCollections(false, true, 0);
            });

            this.setState({ hideUnfollowDialog: true });
            this.props.callBackForLatestFollowedCollections(true);
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _unFollowCollection()", error, "UnFollow Collection");
        }
    }
}