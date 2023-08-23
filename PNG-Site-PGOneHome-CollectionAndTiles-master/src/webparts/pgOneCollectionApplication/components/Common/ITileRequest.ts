export interface ITileRequest {
  Id?: number;
  Title?: string;
  Description?: string;
  SearchKeywords?: string;
  Link?: string;
  OwnerEmail?: string;
  AvailableExternal?: number;
  RequestedById?: number;
  ColorCodeId?: number;
  RequestedAction?: string;
  RequestedDate?: Date;
  ApprovalStatus?: string;
  DecisionDate?: Date;
  DecisionBy?: string;
  DecisionComments?: string;
  ExistingItemID?: number;
}
