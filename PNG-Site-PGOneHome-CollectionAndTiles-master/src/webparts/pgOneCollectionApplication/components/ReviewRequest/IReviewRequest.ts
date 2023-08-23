import { IColorMasterList } from "../Common/IColormasterList";
import { IApplicationList } from "../Common/IApplicationList";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICollectionList } from "../Common/ICollectionList";

export interface IReviewRequestProps {
  context: WebPartContext;
  resourceListItems: any[];
  callBackForRequestSection: any;
}

export interface IReviewRequestState {
  ddlRequestTypeValue: string;
  ddlApprovalStatusValue: string;
  currentRequestType?: string;
  currentRequestStatus?: string;
  detailedListItems?: any[];
  detailedListColumns?: any;
  isRecordAvailable?: boolean;
  isFilterVisible?: boolean;
  isSortedDesc?: boolean;
  lstColorMaster?: IColorMasterList[];
  lstApplicationMaster?: IApplicationList[];
  lstSearchTilesResult?: IApplicationList[];
  bgColor?: string;
  foreColor?: string;
  viewScreenControl?: boolean;
  recordsFound?: number;
  isItemLoaded?: boolean;
  resourceListItems: any[];
  currentPage?: number;
  itemsPerPage?: number;
  totalPage?: number;
  detailedListPageItems?: any[];
  isDataLoaded?: boolean;
  isCorporateCollection?: boolean;

  tNumber?: string;
  tNumberId?: number;
  saveControlsErrorMessage?: string;
  valueRequiredErrorMessage: string;
  tileRequestId?: number;
  tileRequestExistingItemId?: number;
  tileRequestIsPanelOpen?: boolean;
  tileRequestTitle?: string;
  tileRequestDescription?: string;
  tileRequestKeywords?: string;
  tileRequestUrlLink?: string;
  tileRequestIsUrlValid?: boolean;
  tileRequestOwnerEmail?: string;
  tileRequestRequestedAction?: string;
  tileRequestRequestedDate?: string;
  tileRequestColorCode?: number;
  tileRequestDecisionBy?: string;
  tileRequestDecisionDate?: string;
  tileRequestDecisionComments?: string;
  tileRequestShowMessageBar?: boolean;
  tileRequestAvailableExternal?: number;


  colRequestIsPanelOpen?: boolean;
  colRequestShowMessageBar?: boolean;
  colID?: number;
  colTitle?: string;
  //colDefaultMyCollection?: number;
  colDescription?: string;
  colPublicCollection?: number;
  colCorporateCollection?: number;
  colStandardOrder?: number;
  colCollectionOwnerId?: number;
  colCollectionOwnerName?: string;
  colUnDeletable?: number;
  colRequestedDate?: string;
  colApprovalStatus?: string;
  colDecisionDate?: string;
  colDecisionBy?: string;
  colDecisionComments?: string;
  colExistingItemID?: number;
  colRequestedAction?: string;
  colSelectedTiles?: any[];

  colMasterCorporateCollection?: number;
  disableColButtons?: boolean;
}
