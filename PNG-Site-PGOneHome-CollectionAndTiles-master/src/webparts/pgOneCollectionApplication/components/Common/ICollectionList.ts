export interface ICollectionList {
    Title: string;
    ID?: string;
    CollectionOwner?: string;
    Description?: string;
    PublicCollection?: number;
    UnDeletable?: number;
    CollectionOrder?: number;
    DefaultMyCollection?: number;
    CorporateCollection?: number;
    StandardOrder?: number;
    CollectionOwnerId?: number;   
    CollectionOwnerEmail?:string;
}