
export interface IApplicationListState {
    applicationItems: IApplicationList[];
}

export interface IColorCode {
    BgColor?: string;
    ForeColor?: string;
    Title?:string;
}
export interface IApplicationList {
    Title: string;
    ID?: string;
    Link?: string;
    Description?: string;
    OwnerEmail?: string;
    SearchKeywords?: string;
    AvailableExternal?: number;
    ColorCode?: IColorCode;
    ApplicationCollectionMatrixID?: number;
    AppOrder?: number;
    IsActive?: boolean;
}