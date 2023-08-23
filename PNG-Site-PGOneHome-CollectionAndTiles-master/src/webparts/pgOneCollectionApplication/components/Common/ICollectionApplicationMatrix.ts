export interface ICollectionApplicationMatrixState {
    collectionApplicationItems: ICollectionApplicationMatrixList[];
}

export interface ICollectionApplicationMatrixList {
    AppOrder: string;
    ApplicationID: string;
    CollectionID: string;
}