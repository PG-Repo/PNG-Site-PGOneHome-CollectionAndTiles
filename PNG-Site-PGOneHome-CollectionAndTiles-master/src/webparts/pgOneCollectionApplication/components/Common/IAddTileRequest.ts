import { IColorMasterList } from "./IColormasterList";

export interface IAddTileRequest {
  tileRequestIsPanelOpen?: boolean;
  tileRequestTitle?: string;
  tileRequestDescription?: string;
  tileRequestKeywords?: string;
  tileRequestUrlLink?: string;
  tileRequestIsUrlValid?: boolean;
  tileRequestOwnerEmail?: string;
  tileRequestRequestedBy?: string;
  tileRequestShowMessageBar?: boolean;
  tileRequestColorCode?: number;
  lstColorMaster?: IColorMasterList[];
  tileRequestAvailableExternal?: number;
}
