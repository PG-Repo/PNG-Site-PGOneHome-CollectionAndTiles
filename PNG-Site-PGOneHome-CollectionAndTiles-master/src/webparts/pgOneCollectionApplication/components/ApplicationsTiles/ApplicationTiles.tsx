import * as React from 'react';
import { IApplicationList } from '../Common/IApplicationList';
import { arrayMove, SortableContainer, SortableElement } from 'react-sortable-hoc';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import styles from './ApplicationsTiles.module.scss';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { PnPHelper } from '../PnPHelper/PnPHelper';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { gAnalytics } from '../GA/gAnalytics';
import { ICollectionList } from '../Common/ICollectionList';
import { MessageBar, MessageBarType, IContextualMenuItem, Link } from 'office-ui-fabric-react';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { IPoint } from 'office-ui-fabric-react/lib/utilities/positioning/index';
import { HttpClient, IHttpClientOptions } from "@microsoft/sp-http";

export interface IApplicationTilesProps {
    matchedApplicationListBasedOnCollection: IApplicationList[];
    resourceListItems: any[];
    isOwnerForCurrentCollection: boolean;
    webpartContext: WebPartContext;
    callBackForRemovedTilesToBeUpdated: any;
    gAnalytics?: gAnalytics;
    currentCollectionItem?: ICollectionList;
    isAtleastOneVPNTilePresent: boolean;
}

export interface IApplicationTilesState {
    applicationTiles?: IApplicationList[];
    currentApplicationTileItem?: IApplicationList;
    resourceListItems: any[];
    hideRemoveApplicationDialog: boolean;
    isOwnerForCurrentCollection: boolean;
    isShowLoaderSpinnerInModal: boolean;
    matches?: any;
    showVpnCalloutMessage: boolean;
    isAtleastOneVPNTilePresent: boolean;
    isVpnDisconnected: boolean;
    isShowLoaderSpinnerInModalForVPN: boolean;
}

export class ApplicationTiles extends React.Component<IApplicationTilesProps, IApplicationTilesState> {
    private errTitle: string = "PNG-Site-PGOneHome-CollectionAndTiles";
    private errModule: string = "CollectionsLeftNavigation.tsx";
    private pnpHelper: PnPHelper;
    private mouseClickPoint: IPoint;
    private _currentSelectedApplication: IApplicationList = { Title: '', ID: '', Link: '', Description: '', OwnerEmail: '', SearchKeywords: '', ColorCode: { BgColor: '', ForeColor: '', Title: '' } };

    constructor(props: IApplicationTilesProps, state: IApplicationTilesState) {
        super(props);
        this.state = {
            applicationTiles: this.props.matchedApplicationListBasedOnCollection,
            resourceListItems: this.props.resourceListItems,
            hideRemoveApplicationDialog: true,
            isOwnerForCurrentCollection: this.props.isOwnerForCurrentCollection,
            isShowLoaderSpinnerInModal: false,
            matches: window.matchMedia("(min-width: 768px)").matches,
            showVpnCalloutMessage: false,
            isAtleastOneVPNTilePresent: this.props.isAtleastOneVPNTilePresent,
            isVpnDisconnected: true,
            isShowLoaderSpinnerInModalForVPN: false
        };
        this.pnpHelper = new PnPHelper(this.props.webpartContext);
    }

    public componentDidMount() {
        const handler = (e: any) => this.setState({ matches: e.matches });
        window.matchMedia("(min-width: 768px)").addListener(handler);
    }

    private SortableItem = SortableElement(({ value }: { value: IApplicationList }) =>
        <div className={styles.applicationTile}
            onClick={(event) => this._onTileClick(event, value)}
            title={`${value.Title}  :  ${value.Description} ${"\n\n"}URL : ${value.Link}`}            
            style={{ border: "1px solid", borderColor: value.ColorCode.BgColor, backgroundColor: value.ColorCode.BgColor, color: value.ColorCode.ForeColor, position: "relative", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}
        >
            <a href="#" style={{ textDecoration: "none" }} data-interception="off">
                <div className={styles.applicationTileItem} style={{ color: value.ColorCode.ForeColor, textAlign: "center" }}>
                    {this.state.matches ? <span style={{ fontSize: "20px", padding: "3px" }} >{value.Title}</span>
                        : <span style={{ fontSize: "16px" }} >{value.Title}</span>}
                </div>
            </a>
            {this.state.isOwnerForCurrentCollection &&
             <Icon aria-label="Remove tile" role="presentation" style={{ top: "5px", right: "5px", position: "absolute" }} className={`ms-hiddenXlDown ${styles.deleteIcon}`} tabIndex={0} iconName="Delete" title="Remove Tile" 
            onClick={(event) => this._showRemoveApplicationDialog(event, value)} 
            onKeyPress={(event)=> {
                //it triggers by pressing the enter key
              if (event.key === "Enter") {
                this._showRemoveApplicationDialogWithKeyBoard(event, value);
              }
            }}
            />
            }
            {(this.state.matches && value.AvailableExternal == 0) && <Icon role="img" aria-label="Pulse secure VPN Shield" style={{ top: "5px", left: "5px", position: "absolute", fontWeight: "bold", width: "14px" }} className={`${styles.vpnIcon}`} iconName="Shield" title={this.state.resourceListItems["VPN_Shield_Icon_Tooltip"]} />}
            {(!this.state.matches && value.AvailableExternal == 0) && <Icon role="img" aria-label="Pulse secure VPN Shield" style={{ top: "5px", left: "5px", position: "absolute", fontWeight: "bold", width: "14px" }} className={`${styles.vpnIcon}`} iconName="Shield" title={this.state.resourceListItems["VPN_Shield_Icon_Tooltip"]} onClick={(event) => this._showVPNConnectionCallOut(event)} />}
            {this.state.showVpnCalloutMessage &&
                <Callout
                    role="alertdialog"
                    gapSpace={0}
                    setInitialFocus
                    target={this.mouseClickPoint}
                    className={styles.callout}
                    onDismiss={() => { this.setState({ showVpnCalloutMessage: !this.state.showVpnCalloutMessage }); }}
                >
                    <div className={styles.vpnCallOutContent}>
                        <p>{this.state.resourceListItems["VPN_Shield_Icon_Tooltip"]} </p>
                    </div>
                </Callout>
            }

        </div>
    );

    private _onTileClick = async (event: React.MouseEvent<HTMLElement>, value: any) => {
        //Google Analytics: tile Hit in collection (ApplicationTiles)
        this.props.gAnalytics.appHitInCollection(value.Title, value.Link, this.props.currentCollectionItem.Title, this.props.currentCollectionItem.ID, value.ID, false);
        if (value.AvailableExternal == 0) {
            this._currentSelectedApplication = value;
            this.setState({ isShowLoaderSpinnerInModalForVPN: true });
            let configValues = await this.pnpHelper.getConfigMasterListItems();
            const httpClientOptions: IHttpClientOptions = {
                headers: new Headers(),
                method: "GET",
                mode: "cors"
            };
            this.props.webpartContext.httpClient.get(configValues["VPNConnectivityCheckUrl"], HttpClient.configurations.v1, httpClientOptions)
                .then((response) => {
                    if (response.status == 200 || response.status == 202) {
                        window.open(value.Link, '_blank');
                        this.setState({
                            isVpnDisconnected: true,
                            isShowLoaderSpinnerInModalForVPN: false
                        });
                    }
                })
                .catch((error) => {
                    setTimeout(() => {
                        this.setState({
                            isVpnDisconnected: false,
                            isShowLoaderSpinnerInModalForVPN: false
                        });
                    }, 500);
                });
        }
        else {
            window.open(value.Link, '_blank');
        }
    }

    private SortableList = SortableContainer(({ items }: { items: IApplicationList[] }) => {
        return (
            <div className={styles.applicationTilesWrapper}>
                {items.map((value, index) => (
                    <this.SortableItem key={`item-${index}`} index={index} value={value} disabled={!this.state.isOwnerForCurrentCollection} />
                ))}
            </div>
        );
    });
    //Component recieve properties from Parent component
    public componentWillReceiveProps(newProps: IApplicationTilesProps) {
        this.setState({
            applicationTiles: newProps.matchedApplicationListBasedOnCollection,
            resourceListItems: newProps.resourceListItems,
            isOwnerForCurrentCollection: newProps.isOwnerForCurrentCollection,
            isAtleastOneVPNTilePresent: newProps.matchedApplicationListBasedOnCollection.some((tile: any) => {
                return tile.AvailableExternal == 0;
            })
        });
    }
    private onSortEnd = ({ oldIndex, newIndex }: { oldIndex: number, newIndex: number }) => {
        //Google Analytics:tile reorder fired
        this.props.gAnalytics.appReorderFired(this.props.currentCollectionItem.Title, this.props.currentCollectionItem.ID);
        this.setState({
            applicationTiles: arrayMove(this.state.applicationTiles, oldIndex, newIndex),
            isShowLoaderSpinnerInModal: true
        });
        //Update Sort order in CollectionApp Matrix list
        this.pnpHelper.updateSortingOrderInCollectionApplicationMatrixList(this.state.applicationTiles).then(() => {
            this.setState({ isShowLoaderSpinnerInModal: false });
        });
    }

    public render() {
        return (
            <div className={styles.applicationTiles}>
                {this.state.isShowLoaderSpinnerInModal && //show updating Modal on drag and drop
                    <Modal
                        isOpen={true}
                        isBlocking={true}
                        className={styles.spinnerModal}
                    >
                        <Spinner className={styles.showLoaderSpinner} size={SpinnerSize.large} label="Updating..." ariaLive="assertive" labelPosition="right" />
                    </Modal>
                }
                {this.state.isShowLoaderSpinnerInModalForVPN && //show Please Wait.. Checking VPN...
                    <Modal
                        isOpen={true}
                        isBlocking={true}
                        className={styles.spinnerModal}
                    >
                        <Spinner className={styles.showLoaderSpinner} size={SpinnerSize.large} label={this.state.resourceListItems["VPN_Connection_Checking_Message"]} ariaLive="assertive" labelPosition="right" />
                    </Modal>
                }
                {this.state.applicationTiles.length == 0 // show message if zero tile
                    ? <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>{this.state.resourceListItems["emptyCollectionMessage"]}</MessageBar>
                    : <div>
                        {this.state.matches ?
                            <div>
                                {this.state.isAtleastOneVPNTilePresent && <div className={styles.vpnInstruction}><Icon iconName="Shield" /> <span>{this.state.resourceListItems["VPN_Footer_Instruction_Text"]}</span></div>}
                                <this.SortableList items={this.state.applicationTiles} distance={5} onSortEnd={this.onSortEnd} axis="xy" />
                            </div>
                            : <div>
                                {this.state.isAtleastOneVPNTilePresent && <div className={styles.vpnInstruction}><Icon iconName="Shield" /> <span>{this.state.resourceListItems["VPN_Footer_Instruction_Text"]}</span></div>}
                                <this.SortableList items={this.state.applicationTiles} pressDelay={300} onSortEnd={this.onSortEnd} axis="xy" />
                            </div>
                        }
                    </div>
                }

                {this.state.applicationTiles.length > 1
                    ? // confirmation message for deleting the tile
                    <Dialog
                        className={styles.removeApplicationDialog}
                        hidden={this.state.hideRemoveApplicationDialog}
                        onDismiss={this.hideRemoveApplicationDialog.bind(this)}
                        dialogContentProps={{
                            type: DialogType.normal,
                            showCloseButton: false
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 650 } }
                        }}
                    >
                        <div className={styles.removeApplicationDialogContent}>
                            <Icon aria-hidden="true" iconName="Info" />
                            <p className={styles.removeApplicationDialogTitle}>{this.state.resourceListItems["remove_application_Dialog_Title"]}</p>
                            <p className={styles.removeApplicationDialogText}>{this.state.resourceListItems["remove_application_Dialog_Text"]}</p>
                        </div>
                        <DialogFooter>
                            <DefaultButton onClick={this.hideRemoveApplicationDialog.bind(this)} title={this.state.resourceListItems["cancel"]} text={this.state.resourceListItems["cancel"]} />
                            <PrimaryButton className={styles.removeApplicationbutton} onClick={this.removeApplicationTileFromCollection.bind(this)} title={this.state.resourceListItems["remove_application_Dialog_Remove_Button_Text"]} text={this.state.resourceListItems["remove_application_Dialog_Remove_Button_Text"]} />
                        </DialogFooter>
                    </Dialog>
                    : // error message for deleting the last tile
                    <Dialog
                        className={styles.removeLastApplicationDialog}
                        hidden={this.state.hideRemoveApplicationDialog}
                        onDismiss={this.hideRemoveApplicationDialog.bind(this)}
                        dialogContentProps={{
                            type: DialogType.normal,
                            showCloseButton: false
                        }}
                        modalProps={{
                            isBlocking: true,
                            styles: { main: { maxWidth: 650 } }
                        }}
                    >
                        <div className={styles.removeLastApplicationDialogContent}>
                            <Icon aria-hidden="true" iconName="ErrorBadge" />
                            <p className={styles.removeLastApplicationDialogText}>{this.state.resourceListItems["remove_last_application_Dialog_Text"]}</p>
                        </div>
                        <DialogFooter>
                            <DefaultButton onClick={this.hideRemoveApplicationDialog.bind(this)} title={this.state.resourceListItems["cancel"]} text={this.state.resourceListItems["cancel"]} />
                        </DialogFooter>
                    </Dialog>
                }
                <Dialog
                    className={styles.removeLastApplicationDialog}
                    hidden={this.state.isVpnDisconnected}
                    onDismiss={this.hideVpnConnectionDialog.bind(this)}
                    dialogContentProps={{
                        type: DialogType.normal,
                        showCloseButton: false
                    }}
                    modalProps={{
                        isBlocking: true,
                        styles: { main: { maxWidth: 650 } }
                    }}
                >
                    <div className={styles.removeLastApplicationDialogContent}>
                        <Icon aria-hidden="true" iconName="ErrorBadge" />
                        <p className={styles.removeLastApplicationDialogText}>{this.state.resourceListItems["VPN_Disconnected_Error_Message_In_PopUp"]}</p>
                    </div>
                    <DialogFooter className={styles.vpnConnectivityCheckDialogFooter}>
                        <DefaultButton onClick={this.hideVpnConnectionDialog.bind(this)} title={this.state.resourceListItems["VPN_Disconnected_PopUp_button_text"]} text={this.state.resourceListItems["VPN_Disconnected_PopUp_button_text"]} />
                        <Link onClick={this._tryAnywayLinkClick.bind(this)} href={this._currentSelectedApplication.Link} target="_blank" data-interception="off" title={this.state.resourceListItems["VPN_Connection_Checking_Try_Anyway_text"]} >{this.state.resourceListItems["VPN_Connection_Checking_Try_Anyway_text"]}</Link>
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

    private removeApplicationTileFromCollection = (): void => {
        try {
            //Google Analytics:tile reorder fired
            this.props.gAnalytics.appDirectRemove(
                this.state.currentApplicationTileItem.Title,
                this.state.currentApplicationTileItem.Link,
                this.props.currentCollectionItem.Title,
                this.props.currentCollectionItem.ID,
                this.state.currentApplicationTileItem.ID
            );

            this.setState({
                isShowLoaderSpinnerInModal: true,
                hideRemoveApplicationDialog: true
            });
            let currentApplicationTileItem = [];
            currentApplicationTileItem.push(this.state.currentApplicationTileItem);
            // delete the application tile from collection application matrix list
            this.pnpHelper.deleteItemsOnCollectionApplicationMatrixList(currentApplicationTileItem).then(() => {
                let latestApplicationList = this.state.applicationTiles.filter((applicationTile: any) => { return (applicationTile.ID != this.state.currentApplicationTileItem.ID); });
                this.setState({
                    isShowLoaderSpinnerInModal: false,
                    hideRemoveApplicationDialog: true,
                    applicationTiles: latestApplicationList,
                });
                this.props.callBackForRemovedTilesToBeUpdated(latestApplicationList);
            });
        }
        catch (error) {
            this.pnpHelper.errorLogging.logError(this.errTitle, this.errModule, "Some Error Occured in _saveNewCollectionAndApplications()", error, "Create a new Collection");
        }
    }
    // Delete application tile pop up
    private _showRemoveApplicationDialog = (event: React.MouseEvent<HTMLElement>, currentApplicationTileItem: IApplicationList) => {
        event.stopPropagation();
        this.setState({
            hideRemoveApplicationDialog: false,
            currentApplicationTileItem: currentApplicationTileItem
        });
    }
    // Delete application tile pop up with Keyboard
    private _showRemoveApplicationDialogWithKeyBoard = (event: React.KeyboardEvent<HTMLElement>, currentApplicationTileItem: IApplicationList) => {
        event.stopPropagation();
        this.setState({
            hideRemoveApplicationDialog: false,
            currentApplicationTileItem: currentApplicationTileItem
        });
    }

    // hide dialog
    private hideRemoveApplicationDialog = (): void => {
        this.setState({ hideRemoveApplicationDialog: true });
    }

    private hideVpnConnectionDialog = (): void => {
        this.setState({ isVpnDisconnected: true });
    }

    private _tryAnywayLinkClick = (): void => {
        //Google Analytics:Try Link Clicked
        this.props.gAnalytics.appTryLinkClicked();
        this.setState({ isVpnDisconnected: true });
    }
    //open VPN Connectivity call out
    private _showVPNConnectionCallOut = (ev?: any, item?: IContextualMenuItem): boolean | void => {
        this.mouseClickPoint = { x: ev.clientX - 3, y: ev.clientY + 3 };
        ev.stopPropagation();
        this.setState({
            showVpnCalloutMessage: !this.state.showVpnCalloutMessage
        });
    }
}
