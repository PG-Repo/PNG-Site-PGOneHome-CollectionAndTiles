import * as React from 'react';
import styles from './PgOneCollectionApplication.module.scss';
import { IPgOneCollectionApplicationProps } from './IPgOneCollectionApplicationProps';
import { CollectionsLeftNavigation } from './CollectionsLeftNavigation/CollectionsLeftNavigation';
import { CollectionTopSettings } from "./CollectionTopSettings/CollectionTopSettings";
import { ApplicationTiles } from "./ApplicationsTiles/ApplicationTiles";
import { ICollectionList } from "./Common/ICollectionList";
import { IApplicationList } from './Common/IApplicationList';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { PnPHelper } from './PnPHelper/PnPHelper';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { MessageBar, MessageBarType, Link, ContextualMenu } from 'office-ui-fabric-react';
import { ReviewRequest } from '../components/ReviewRequest/ReviewRequest';
import { ManageMyTiles } from "./ManageMyTiles/ManageMyTiles";
import { ManageSnowRequests } from "./ManageSnowRequest/ManageSnowRequests";
import { TrackRequest } from "./TrackRequest/TrackRequest";
import { gAnalytics } from './GA/gAnalytics';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

export interface IPgOneCollectionApplicationState {
  loading: boolean;
  currentCollectionItem: ICollectionList;
  myFollowedCollectionsList: ICollectionList[];
  matchedApplicationListBasedOnCollection: IApplicationList[];
  currentFollowedCollection?: ICollectionList[];
  isCurrentCollectionBeingFollowed: boolean;
  isValidCollectionID: boolean;
  resourceListItem?: any[];
  showRequestSection: boolean;
  showReviewRequestSection: boolean;
  showTrackRequestSection: boolean;
  currentSectionTitle?: string;
  showLeftNavigationInSmallScreen: boolean;
  showSuccessMessageForPublicCollectionRequest: boolean;
  showErrorMessage: boolean;
  errorMessageContent: string;
  isOwnerForCurrentCollection: boolean;
  successMessageForPublicCollectionRequest?: string;
  isSettingMenuExpanded: boolean;
  loadReadonlyUserProfile: boolean;
  dissableReSetButton:boolean;
}

export default class PgOneCollectionApplication extends React.Component<IPgOneCollectionApplicationProps, IPgOneCollectionApplicationState> {

  private pnpHelper: PnPHelper;
  private gAnalytics: gAnalytics;
  private queryParms = new UrlQueryParameterCollection(window.location.href);
  private currentCollectionID = this.queryParms.getValue("COLLECTIONID");
  private pageView = this.queryParms.getValue("VIEW");
  private isValidCollectionID: boolean;
  private currentCollectionDetails: ICollectionList;
  private currentUserTNumber: string = "";
  private WebPartName: string = "PNG-Site-PGOneHome-CollectionAndTiles";
  private Module: string = "PgOneCollectionApplication.tsx";

  constructor(props: IPgOneCollectionApplicationProps, state: IPgOneCollectionApplicationState) {
    super(props);
    this.state = {
      loading: true,
      currentCollectionItem: { Title: '', ID: '', CollectionOwner: '', Description: '', PublicCollection: 0, UnDeletable: 0, CollectionOrder: 0 },
      myFollowedCollectionsList: [],
      matchedApplicationListBasedOnCollection: [],
      isCurrentCollectionBeingFollowed: true,
      isValidCollectionID: false,
      showRequestSection: false,
      showReviewRequestSection: false,
      showTrackRequestSection: false,
      showLeftNavigationInSmallScreen: false,
      showSuccessMessageForPublicCollectionRequest: false,
      showErrorMessage: false,
      errorMessageContent: '',
      isOwnerForCurrentCollection: false,
      successMessageForPublicCollectionRequest: "",
      isSettingMenuExpanded: false,
      loadReadonlyUserProfile: false,
      dissableReSetButton:false,
    };

    this.pnpHelper = new PnPHelper(this.props.webPartContext);

  }

  public componentWillMount() {
    try {
      Promise.all([
        this.pnpHelper.getResourceListItems(),
        this.pnpHelper.getCurrentUserCollections(),
        this.pnpHelper.userProps("TNumber"),
        this.pnpHelper.getConfigMasterListItems()
      ]).then(async ([resourceListItems, myFollowedCollectionsList, currentUserTNumber, configValues]) => {

        //Config values
        let trackingId = configValues['GoogleAnalyticsTrackingId'];
        let currentUser = currentUserTNumber;
        this.gAnalytics = new gAnalytics(trackingId, currentUser);
        this.pageView = currentUserTNumber == configValues["DefaultReadOnlyUserProfileName"] ? undefined : this.pageView;
        // To check for Request screen landing
        if (this.pageView != undefined) {
          this.setState({
            loading: false,
            isSettingMenuExpanded: true,
            myFollowedCollectionsList: myFollowedCollectionsList,
            showErrorMessage: false,
            resourceListItem: resourceListItems,
            showRequestSection: true,
            currentCollectionItem: { Title: '', ID: '', CollectionOwner: '', Description: '', PublicCollection: 0, UnDeletable: 0, CollectionOrder: 0 },
            currentSectionTitle: this.pageView.toLocaleLowerCase() == "trackrequests" ? "Track requests" : this.pageView.toLocaleLowerCase() == "reviewrequests" ? "Review requests" : "",
          });
        }
        // To check any Collection ID is present in Query string
        else if (this.currentCollectionID != undefined) {
          if (isNaN(Number(this.currentCollectionID))) {
            this.isValidCollectionID = false;
          }
          else {
            this.currentCollectionDetails = await this.pnpHelper.getCurrentCollectionDetails(this.currentCollectionID, myFollowedCollectionsList);
            this.isValidCollectionID = this.currentCollectionDetails != undefined ? true : false;

            //Google Analytics: validation for Colection id
            if (this.isValidCollectionID) {
              //Google Analytics: Collection Hit from Querystring?
              this.gAnalytics.collectionHitFromGet(this.currentCollectionDetails.Title, this.currentCollectionDetails.ID);
            }
          }
        }
        this.currentUserTNumber = currentUserTNumber.toLocaleLowerCase();
        let currentCollectionItem = this.currentCollectionID != undefined ? this.currentCollectionDetails : myFollowedCollectionsList[0];
        if (currentCollectionItem != undefined) {
          // to get current collection application list
          this.pnpHelper.getCurrentCollectionApplications(currentCollectionItem.ID).then((matchedApplicationListBasedOnCollection: any) => {
            this.setState({
              loading: false,
              showErrorMessage: false,
              loadReadonlyUserProfile: currentUserTNumber == configValues["DefaultReadOnlyUserProfileName"],              
              resourceListItem: resourceListItems,
              myFollowedCollectionsList: myFollowedCollectionsList,
              currentCollectionItem: (this.pageView !== undefined ? { Title: '', ID: '', CollectionOwner: '', Description: '', PublicCollection: 0, UnDeletable: 0, CollectionOrder: 0 } : currentCollectionItem),
              isCurrentCollectionBeingFollowed: this.currentCollectionID != undefined ? myFollowedCollectionsList.some((collectionItem: any) => { return String(collectionItem.ID) === this.currentCollectionID; }) : true,
              isOwnerForCurrentCollection: currentCollectionItem.CollectionOwner.toLocaleLowerCase() == this.currentUserTNumber,
              isValidCollectionID: true,
              matchedApplicationListBasedOnCollection: matchedApplicationListBasedOnCollection.sort((a: any, b: any) => { return a.AppOrder - b.AppOrder; }),
            });
          });
        }
        else {
          this.setState({
            loading: false,
            myFollowedCollectionsList: myFollowedCollectionsList,
            resourceListItem: resourceListItems,
            showRequestSection: false,
            isValidCollectionID: false
          });
        }
      })
        .catch((err: any) => {
          this.setState({
            loading: false,
            showErrorMessage: true,
            errorMessageContent: err
          });
        });
    }
    catch (error) {
      this.pnpHelper.errorLogging.logError(this.WebPartName, this.Module, "Some Error Occured in componentWillMount()", error, "Page On Load");
    }
  }

  public render(): React.ReactElement<IPgOneCollectionApplicationProps> {
    return (
      <div className="ms-Grid">
        {this.state.loading
          ? <div className={`ms-Grid-row`}>
            <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12 ${styles.spinner}`}>
              <Spinner size={SpinnerSize.medium} label="Loading..." ariaLive="assertive" labelPosition="right" />
            </div>
          </div>
          : <div>            
            {this.state.showErrorMessage
              ? <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                <MessageBar className={styles.errorMessageBar} messageBarType={MessageBarType.warning} isMultiline={false}>{this.state.errorMessageContent}</MessageBar>
              </div>
              : <div>
                <div className={`ms-Grid-row ${styles.pgOneCollectionApplication}`}>
                  {this.state.loadReadonlyUserProfile &&
                    <div className={`ms-Grid-col ms-sm12 ms-md12 ms-lg12`}>
                      <MessageBar className={styles.readOnlyUserMessageBar} messageBarType={MessageBarType.severeWarning} isMultiline={false}>{this.state.resourceListItem["readOnlyUserMessage"]}
                        <Link target="_blank" data-interception="off" href={this.state.resourceListItem["BrowserWarningLinkUrl"]}>{this.state.resourceListItem["BrowserWarningLinkDisplayText"]}</Link>.</MessageBar>
                    </div>}
                  <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg3 ${this.state.showLeftNavigationInSmallScreen ? `${styles.showLeftNavigation}` : `${styles.pgOneCollectionsLeftNavigation}`} `}>
                    {/* Visible Only on smaller screen */}
                    <div className={`ms-Grid-col ms-sm12 ms-hiddenLgUp ${styles.cancelLeftNavigation}`}>
                      <Icon tabIndex={1} onClick={this._showLeftNavigationInSmallScreen.bind(this)} iconName="Cancel" />
                    </div>
                    <CollectionsLeftNavigation
                      myFollowedCollectionsList={this.state.myFollowedCollectionsList}
                      callBackHandlerForTopSettings={this._currentCollectionItemUpdateHandler.bind(this)}
                      webpartContext={this.props.webPartContext}
                      currentActiveCollection={this.state.currentCollectionItem}
                      resourceListItems={this.state.resourceListItem}
                      callBackHandlerForTrackRequests={this._showTrackRequestsHandler.bind(this)}
                      callBackHandlerForReviewRequests={this._showReviewRequestsHandler.bind(this)}
                      currentSectionTitle={this.state.currentSectionTitle}
                      callBackForLatestFollowedCollections={this._showLoadingAndRefreshResult.bind(this)}
                      gAnalytics={this.gAnalytics}
                      isSettingMenuExpanded={this.state.isSettingMenuExpanded}
                      loadReadonlyUserProfile={this.state.loadReadonlyUserProfile}
                    />
                  </div>
                  <div className={`ms-Grid-col ms-sm12 ms-md6 ms-lg9 ${styles.pgOneCollectionsMiddleContent}`}>
                    {/* Visible Only on smaller screen */}
                    <div className={`ms-Grid-col ms-sm1 ms-md6 ms-hiddenLgUp ${styles.openLeftNavigationMenu}`}>
                      <Icon onClick={this._showLeftNavigationInSmallScreen.bind(this)} iconName="CollapseMenu" />
                    </div>
                    {this.state.showSuccessMessageForPublicCollectionRequest &&
                      <div className={`ms-Grid-col ms-sm10 ms-md6 ms-lg12 ${styles.successMessageBar}`}>
                        <MessageBar messageBarType={MessageBarType.success} isMultiline={false}
                          onDismiss={() => { this.setState({ showSuccessMessageForPublicCollectionRequest: false }); }}
                          dismissButtonAriaLabel="Close">
                          {this.state.successMessageForPublicCollectionRequest}
                        </MessageBar>
                      </div>
                    }
                    {!this.state.showRequestSection
                      ? <div>
                        {this.state.isValidCollectionID
                          ?
                          // Collection top settings Section & applicationTiles
                          <div>
                            <CollectionTopSettings
                              currentCollectionItem={this.state.currentCollectionItem}
                              myFollowedCollectionsList={this.state.myFollowedCollectionsList}
                              webpartContext={this.props.webPartContext}
                              matchedApplicationListBasedOnCollection={this.state.matchedApplicationListBasedOnCollection}
                              isCurrentCollectionBeingFollowed={this.state.isCurrentCollectionBeingFollowed}
                              resourceListItems={this.state.resourceListItem}
                              callBackForLatestFollowedCollections={this._showLoadingAndRefreshResult.bind(this)}
                              isOwnerForCurrentCollection={this.state.isOwnerForCurrentCollection}
                              gAnalytics={this.gAnalytics}
                              loadReadonlyUserProfile={this.state.loadReadonlyUserProfile}
                            />
                            <ApplicationTiles
                              matchedApplicationListBasedOnCollection={this.state.matchedApplicationListBasedOnCollection}
                              resourceListItems={this.state.resourceListItem}
                              isOwnerForCurrentCollection={this.state.isOwnerForCurrentCollection}
                              webpartContext={this.props.webPartContext}
                              callBackForRemovedTilesToBeUpdated={this._updateLatestApplicationTilesInCurrentCollection.bind(this)}
                              gAnalytics={this.gAnalytics}
                              currentCollectionItem={this.state.currentCollectionItem}
                              isAtleastOneVPNTilePresent={this.state.matchedApplicationListBasedOnCollection.some((tile: any) => { return tile.AvailableExternal == 0; })}
                            />
                            
                          </div>
                          : <div className={`ms-Grid-col ms-sm10 ms-md6 ms-lg12`}>
                            {this.state.matchedApplicationListBasedOnCollection.length == 0 && this.state.myFollowedCollectionsList.length == 0 || this.state.matchedApplicationListBasedOnCollection.length==0
                              ? 
                              <React.Fragment> 
                                {localStorage.setItem("isFirstTimeUser", "0")}
                                 
                                <MessageBar messageBarType={MessageBarType.warning} isMultiline={true}>
                                {this.state.resourceListItem["SetupMyProfileMessage"]==null
                                ?
                                'Welcome, It Seems you are visiting this site first time. Please click on "Setup My Profile" to generate your new profile which will take 2-3 minutes to complete.'
                                :this.state.resourceListItem["SetupMyProfileMessage"]}
                                  
                                   <Link ref="btnReset" disabled={this.state.dissableReSetButton} onClick={this.reSetProfile.bind(this)} title={this.state.resourceListItem["SetupMyProfileMessageButton"]==null?"Setup My Profile":this.state.resourceListItem["SetupMyProfileMessageButton"]}  >{this.state.resourceListItem["SetupMyProfileMessageButton"]==null?"Setup My Profile":this.state.resourceListItem["SetupMyProfileMessageButton"]} <Icon iconName="ConfigurationSolid" />
                                  </Link>
                                  </MessageBar>
                              </React.Fragment>

                              : <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>{this.state.resourceListItem["invalid_collectionID_message"]}</MessageBar>
                            }
                          </div>
                        }
                      </div>
                      // Requests Section
                      :
                      this.state.currentSectionTitle === "Review requests" ? (
                        <ReviewRequest
                          context={this.props.webPartContext}
                          resourceListItems={this.state.resourceListItem}
                          callBackForRequestSection={this._showSuccessMessageForRequestSection.bind(this)}
                        ></ReviewRequest>
                      ) : this.state.currentSectionTitle === "Track requests" ? (
                        <TrackRequest
                          wpContext={this.props.webPartContext}
                          resourceListItems={this.state.resourceListItem}
                          callBackForRequestSection={this._showSuccessMessageForRequestSection.bind(this)}
                        />
                      ) : this.state.currentSectionTitle === "SNOW requests" ? (
                        <ManageSnowRequests
                          wpContext={this.props.webPartContext}
                          resourceListItems={this.state.resourceListItem}
                          callBackForRequestSection={this._showSuccessMessageForRequestSection.bind(this)}
                        />
                      ): this.state.currentSectionTitle === "Manage my tiles" && (
                        <ManageMyTiles
                          wpContext={this.props.webPartContext}
                          resourceListItems={this.state.resourceListItem}
                          callBackForRequestSection={this._showSuccessMessageForRequestSection.bind(this)}
                        />
                      )
                    }
                  </div>
                </div>
              </div>
            }
          </div>
        }
      </div>
    );
  }
  // call back method for left navigation item click
  private _currentCollectionItemUpdateHandler(item: any, myFollowedCollectionItems: ICollectionList[]) {
    this.pnpHelper.getCurrentCollectionApplications(item.ID).then((matchedApplicationListBasedOnCollection: any) => {
      this.setState({
        currentCollectionItem: item,
        isOwnerForCurrentCollection: item.CollectionOwner.toLocaleLowerCase() == this.currentUserTNumber,
        matchedApplicationListBasedOnCollection: matchedApplicationListBasedOnCollection.sort((a: any, b: any) => { return a.AppOrder - b.AppOrder; }),
        myFollowedCollectionsList: myFollowedCollectionItems,
        isCurrentCollectionBeingFollowed: true,
        showRequestSection: false,
        currentSectionTitle: "",
        isValidCollectionID: true,
        showLeftNavigationInSmallScreen: false,
        isSettingMenuExpanded: false
      });
    });
  }

  private _showTrackRequestsHandler(requestType: string, myFollowedCollectionItems: ICollectionList[]) {
    this.setState({
      currentCollectionItem: { Title: '', ID: '', CollectionOwner: '', Description: '', PublicCollection: 0, UnDeletable: 0, CollectionOrder: 0 },
      currentSectionTitle: requestType,
      showRequestSection: true,
      myFollowedCollectionsList: myFollowedCollectionItems,
      showLeftNavigationInSmallScreen: false,
      isSettingMenuExpanded: true
    });
  }

  private _showReviewRequestsHandler(requestType: string, myFollowedCollectionItems: ICollectionList[]) {
    this.setState({
      currentCollectionItem: { Title: '', ID: '', CollectionOwner: '', Description: '', PublicCollection: 0, UnDeletable: 0, CollectionOrder: 0 },
      currentSectionTitle: requestType,
      showRequestSection: true,
      myFollowedCollectionsList: myFollowedCollectionItems,
      showLeftNavigationInSmallScreen: false,
      isSettingMenuExpanded: true
    });
  }

  private _showLeftNavigationInSmallScreen = (): void => {
    this.setState({
      showLeftNavigationInSmallScreen: !this.state.showLeftNavigationInSmallScreen
    });
  }

  private _updateLatestApplicationTilesInCurrentCollection = (latestApplicationList: any): void => {
    this.setState({
      matchedApplicationListBasedOnCollection: latestApplicationList.sort((a: any, b: any) => { return a.AppOrder - b.AppOrder; }),
    });
  }
  private async _showSuccessMessageForRequestSection(successMessageForRequest?: string) {
    this.setState({
      showSuccessMessageForPublicCollectionRequest: true,
      successMessageForPublicCollectionRequest: successMessageForRequest
    });
    let configValues = await this.pnpHelper.getConfigMasterListItems();
    setTimeout(() => {
      this.setState({ showSuccessMessageForPublicCollectionRequest: false });
    }, parseInt(configValues["SuccessMessageDisplayTime"]));
  }

  //Show Loader and call back method on page load.
  private async _showLoadingAndRefreshResult(isLoading: boolean, isRecallRequired?: boolean, currentItemindex?: number, successMessageForPublicCollectionRequest?: string) {
    if (isLoading) {
      this.setState({
        loading: isLoading
      });
    }
    else {
      if (isRecallRequired) {
        //Query from SharePoint List again
        this.pnpHelper.getCurrentUserCollections().then((myFollowedCollectionsList) => {
          let currentCollectionItem = myFollowedCollectionsList[currentItemindex];
          this.pnpHelper.getCurrentCollectionApplications(currentCollectionItem.ID).then((matchedApplicationListBasedOnCollection: any) => {
            this.setState({
              loading: isLoading,
              myFollowedCollectionsList: myFollowedCollectionsList,
              matchedApplicationListBasedOnCollection: matchedApplicationListBasedOnCollection.sort((a: any, b: any) => { return a.AppOrder - b.AppOrder; }),
              currentCollectionItem: currentCollectionItem,
              isOwnerForCurrentCollection: currentCollectionItem.CollectionOwner.toLocaleLowerCase() == this.currentUserTNumber,
              showLeftNavigationInSmallScreen: false,
              isCurrentCollectionBeingFollowed: true,
              showRequestSection: false,
              currentSectionTitle: "",
              isSettingMenuExpanded: false
            });
          });
        });
      }
      //render items from the already loaded in the page
      if (!isRecallRequired) {
        if (this.state.showRequestSection) {
          this.setState({
            loading: isLoading,
            showLeftNavigationInSmallScreen: false,
            showSuccessMessageForPublicCollectionRequest: true,
            successMessageForPublicCollectionRequest: successMessageForPublicCollectionRequest
          });
        }
        else {
          this.setState({
            loading: isLoading,
            currentCollectionItem: this.state.myFollowedCollectionsList[currentItemindex],
            isOwnerForCurrentCollection: this.state.myFollowedCollectionsList[currentItemindex].CollectionOwner.toLocaleLowerCase() == this.currentUserTNumber,
            showLeftNavigationInSmallScreen: false,
            showSuccessMessageForPublicCollectionRequest: true,
            successMessageForPublicCollectionRequest: successMessageForPublicCollectionRequest,
            showRequestSection: false,
            currentSectionTitle: "",
            isSettingMenuExpanded: false
          });
        }
        // display success message for 8 seconds only
        let configValues = await this.pnpHelper.getConfigMasterListItems();
        setTimeout(() => {
          this.setState({ showSuccessMessageForPublicCollectionRequest: false });
        }, parseInt(configValues["SuccessMessageDisplayTime"]));
      }
    }
  }

  private reSetProfile(){
    //DX2675 is pgonesupport.im@pg.com
    try {
      this.setState({dissableReSetButton:true});
      if(this.currentUserTNumber != ("DX2675").toLocaleLowerCase()){
        Promise.all([
          this.pnpHelper.reSetDefaultProfile(this.currentUserTNumber)
        ]).then(([result]) => {
          localStorage.clear();
          location.reload();
        });
      }
    }
    catch (error) {
      localStorage.clear();
      location.reload();
    }
  }

}
