export interface ICollectionRequest {
  ID?: number;
  Title?: string;
  DefaultMyCollection?: number;
  Description?: string;
  PublicCollection?: number;
  CorporateCollection?: number;
  StandardOrder?: number;
  CollectionOwnerId?: number;
  UnDeletable?: number;
  RequestedDate?: Date;
  ApprovalStatus?: string;
  DecisionDate?: Date;
  DecisionBy?: string;
  DecisionComments?: string;
  ExistingItemID?: number;
  RequestedAction?: string;
}
