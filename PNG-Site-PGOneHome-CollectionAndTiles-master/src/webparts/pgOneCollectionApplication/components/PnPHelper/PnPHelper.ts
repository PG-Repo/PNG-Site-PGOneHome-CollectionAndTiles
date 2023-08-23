import "@pnp/polyfill-ie11";
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import { IItemAddResult, Items } from "@pnp/sp/items";
import { dateAdd } from "@pnp/common";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICollectionList } from "../Common/ICollectionList";
import { IApplicationList } from '../Common/IApplicationList';
import "@pnp/sp/profiles";
import { ErrorLogging } from "../ErrorLogging/ErrorLogging";
import { ITileRequest } from "../Common/ITileRequest";
import { MSGraphClient } from '@microsoft/sp-http';

let UserMasterListName = "UserMaster";
let CollectionMasterListName = "CollectionMaster";
let ApplicationMasterListName = "ApplicationMaster";
let UserCollectionMatrixListName = "UserCollectionMatrix";
let CollectionApplicationMatrixListName = "CollectionApplicationMatrix";
let CollectionRequestsListName = "CollectionRequests";
let CollectionApplicationMatrixRequestsListName = "CollectionApplicationMatrixRequests";
let ConfigMasterListName = "ConfigMaster";
let ResourcesMasterListName = "ResourcesMaster";
let ApplicationRequestsListName = "ApplicationRequests";
let WebPartName = "PNG-Site-PGOneHome-CollectionAndTiles";
let classFileName = "PnPHelper.ts";

export class PnPHelper {
    private webPartContext: WebPartContext;
    private configValues: any;
    private currentUserName: string;
    private siteName: string;
    public errorLogging: ErrorLogging;

    constructor(wpContext: any) {
        sp.setup({
            sp: {
                headers: {
                    "Accept": "application/json; odata=verbose",
                    "ContentType": "application/json; odata=verbose",
                    "User-Agent": "NONISV|PNGOne|PGOneHome/1.0",
                    "X-ClientService-ClientTag": "NONISV|PNGOne|PGOneHome/1.0"
                }
            },
            // set ie 11 mode
            ie11: true,
            spfxContext: wpContext,
        });
        this.webPartContext = wpContext;
        this.currentUserName = this.webPartContext.pageContext.user.loginName;
        this.siteName = this.webPartContext.pageContext.web.title.toLocaleUpperCase();
        this.errorLogging = new ErrorLogging(this.webPartContext);
    }

    //To Get Resourece List Items
    public async getResourceListItems(): Promise<any> {
        let resourceItems = new Array();
        return new Promise<any>(async (resolve: (item: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                this.configValues = await this.getConfigMasterListItems();
                sp.web.lists.getByTitle(ResourcesMasterListName)
                    .items
                    .select("Title", "ValueForKey")
                    .filter("Locale eq 'en'")
                    .top(5000)
                    .usingCaching({
                        expiration: dateAdd(new Date(), "day", parseInt(this.configValues["ResourceListLabelValuesCacheExpiry"])),
                        key: this.siteName + "-Home-Resources",
                        storeName: "local"
                    })
                    .get()
                    .then(async (items: any): Promise<void> => {
                        try {
                            if (items.length > 0) {
                                items.map((value: any, index: any) => {
                                    resourceItems[value["Title"]] = value["ValueForKey"];
                                });
                                resolve(resourceItems);
                            }
                        }
                        catch (error) {
                            reject(error);
                            this.errorLogging.logError(WebPartName, classFileName, "Items in getResourceListItems()", error, "Page On Load");
                        }
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getResourceListItems() while connecting to SharePoint List", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "getResourceListItems()", error, "Page On Load");
            }
        });
    }

    //To Get Application Master List Items
    public async getApplicationTiles(): Promise<IApplicationList[]> {
        return new Promise<IApplicationList[]>(async (resolve, reject) => {
            try {
                this.configValues = await this.getConfigMasterListItems();
                sp.web.lists.getByTitle(ApplicationMasterListName)
                    .items
                    .select("ID,Title,Description,OwnerEmail,Link,SearchKeywords,AvailableExternal,ColorCode/Title,ColorCode/BgColor,ColorCode/ForeColor")
                    .expand("ColorCode")
                    .filter("IsActive eq 1")
                    .getAll()
                    .then((result: any): void => {
                        if (result.length > 0) {
                            result = result.sort((a: any, b: any) => {
                                var x = a.Title.toLowerCase();
                                var y = b.Title.toLowerCase();
                                return (x < y) ? -1 : (x > y) ? 1 : 0;
                            });
                        }
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in getApplicationTiles()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occurred in getApplicationTiles()", error, "Page On Load");
            }
        });
    }

    //To Get Active Public Collections from Collection Master List Items
    public async getPublicCollections(): Promise<ICollectionList[]> {
        var publicCollectionItems: ICollectionList[] = [];
        return new Promise<ICollectionList[]>(async (resolve, reject) => {
            try {
                sp.web.lists.getByTitle(CollectionMasterListName)
                    .items
                    .orderBy("Title")
                    .select("ID,Title,Description,PublicCollection,UnDeletable,CollectionOwner/Title,CollectionOwner/Email,IsActive")
                    .expand("CollectionOwner")
                    .filter("(PublicCollection eq 1) and (IsActive eq 1)")
                    .getAll()
                    .then((result: any): void => {
                        if (result.length > 0) {
                            result.map((value: any) => {
                                value["CollectionOwnerEmail"] = value["CollectionOwner"]["Email"];
                                value["CollectionOwner"] = value["CollectionOwner"]["Title"];
                                publicCollectionItems.push(value);
                            });
                            resolve(publicCollectionItems.sort((a: any, b: any) => {
                                var x = a.Title.toLowerCase();
                                var y = b.Title.toLowerCase();
                                return (x < y) ? -1 : (x > y) ? 1 : 0;
                            }));
                        }
                        else {
                            resolve(publicCollectionItems);
                        }
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in getPublicCollections()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getPublicCollections()", error, "Page On Load");
            }
        });
    }

    //To Get current user collections
    public async getCurrentUserCollections(): Promise<ICollectionList[]> {
        var userCollectionItems: ICollectionList[] = [];
        let userTNumber = "";
        return new Promise<ICollectionList[]>(async (resolve, reject) => {
            try {
                userTNumber = await this.userProps("TNumber");
                if (userTNumber === "ExternalUser") {
                    throw ("Your details is not available with us to use this application. Please contact support help desk.");
                }
                if (userTNumber === this.configValues["DefaultReadOnlyUserProfileName"]) {
                    this.errorLogging.logError(WebPartName, classFileName, "Items in graphUserPropsForTNumber()", "Your TNumber is currently unavailable to access this application. Please contact the support help desk.", "Page On Load");
                }
                //userTNumber = "tf7401";
                this.getCurrentUserItemID(userTNumber)
                    .then(async (userItem: any) => {
                        this._getCurrentUserCollectionIDs(userItem)
                            .then(async (collectionItems: any) => {
                                if (collectionItems.length > 0) {
                                    sp.web.lists.getByTitle(CollectionMasterListName)
                                        .items
                                        .orderBy("ID")
                                        .select("ID,Title,Description,PublicCollection,DefaultMyCollection,CorporateCollection,StandardOrder,UnDeletable,CollectionOwner/Title,IsActive")
                                        .expand("CollectionOwner")
                                        .filter(this._generateFilterCondition("ID", "CollectionID", collectionItems) + " and (IsActive eq 1)")
                                        .getAll()
                                        .then((result: any): void => {
                                            result.map((value: any) => {
                                                var matchedCollectionItem = collectionItems.filter((collectionItem: any) => {
                                                    return (collectionItem["CollectionID"]["ID"] == value["ID"]);
                                                });
                                                if (matchedCollectionItem.length > 0) {
                                                    value["CollectionOrder"] = matchedCollectionItem[0]["CollectionOrder"];
                                                    value["CollectionOwner"] = value["CollectionOwner"]["Title"];
                                                    value["UserCollectionMatrixItemID"] = matchedCollectionItem[0]["ID"];
                                                    userCollectionItems.push(value);
                                                }
                                                else {
                                                    value["CollectionOrder"] = 0;
                                                    userCollectionItems.push(value);
                                                }
                                            });
                                            resolve(userCollectionItems.sort((a: any, b: any) => { return a.CollectionOrder - b.CollectionOrder; }));
                                        }, (error: any): void => {
                                            reject(error);
                                            this.errorLogging.logError(WebPartName, classFileName, "Items in getCurrentUserCollections()", error, "Page On Load");
                                        });

                                } else {
                                    resolve(userCollectionItems);
                                }
                            });
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getCurrentUserCollections()", error, "Page On Load");
            }
        });
    }

    //generate filter condition for multiple collection ids
    private _generateFilterCondition(fieldName: string, lookUpFieldName: string, values: any) {
        let _filterConditions = "";
        if (values.length > 0) {
            values.map((value: any) => {
                _filterConditions += fieldName + ' eq ' + value[lookUpFieldName][fieldName] + ' or ';
            });
        }
        return '(' + _filterConditions.slice(0, -4) + ')';
    }

    // to get user TNumber
    public userProps(userProfileProperty: any): Promise<string> {
        return new Promise<string>(async (resolve: (userProfilePropertyValue: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                this.configValues = await this.getConfigMasterListItems();
                sp.profiles.myProperties
                    .usingCaching({
                        expiration: dateAdd(new Date(), "day", parseInt(this.configValues["UserProfilePropertyCacheExpiry"])),
                        key: this.siteName + "-UserProfile-" + this.currentUserName,
                        storeName: "local"
                    })
                    .get()
                    .then((result: any): void => {
                        if (result.AccountName.match("#ext#")) {
                            resolve("ExternalUser");
                        } else {
                            result["UserProfileProperties"]["results"].map(async (v: any) => {
                                if (v.Key === userProfileProperty) {
                                    if (userProfileProperty === "TNumber") {
                                        // v.Value = ""; // simulating empty TNumber user
                                        if (v.Value == null || v.Value == "") {
                                            // resolve(this.configValues["DefaultReadOnlyUserProfileName"]);
                                            let userTNumber = await this.graphUserPropsForTNumber();
                                            resolve(userTNumber);
                                        }
                                        else {
                                            resolve(v.Value);
                                        }
                                    }
                                    else {
                                        resolve(v.Value);
                                    }
                                }
                            });
                        }
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in userProps()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in userProps()", error, "Page On Load");
            }
        });
    }


    // to get user TNumber
    public graphUserPropsForTNumber(): Promise<string> {
        return new Promise<string>(async (resolve: (userProfilePropertyValue: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                this.configValues = await this.getConfigMasterListItems();
                var itemsGraphUserProfile = this.getLocalStorage("GraphUserProfile");
                if (itemsGraphUserProfile == null) {
                    this.webPartContext.msGraphClientFactory
                        .getClient()
                        .then((client: MSGraphClient): void => {
                            client
                                .api('/me?$select=employeeId,userPrincipalName')
                                .get((error, response: any, rawResponse?: any) => {
                                    // response.employeeId = "";// simulating empty TNumber user
                                    this.setLocalStorage("GraphUserProfile", response, parseInt(this.configValues["UserProfilePropertyCacheExpiry"]));
                                    if (response != null && (response.employeeId != null && response.employeeId != "")) {
                                        resolve(response.employeeId.toLocaleLowerCase());
                                    }
                                    else {
                                        resolve(this.configValues["DefaultReadOnlyUserProfileName"]);
                                    }
                                });
                        }, (error: any): void => {
                            reject(error);
                            this.errorLogging.logError(WebPartName, classFileName, "Items in graphUserPropsForTNumber()", error, "Page On Load");
                        });
                }
                else {
                    if (itemsGraphUserProfile != null && (itemsGraphUserProfile.employeeId != null && itemsGraphUserProfile.employeeId != "")) {
                        resolve(itemsGraphUserProfile.employeeId.toLocaleLowerCase());
                    }
                    else {
                        resolve(this.configValues["DefaultReadOnlyUserProfileName"]);
                    }
                }
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in graphUserPropsForTNumber()", error, "Page On Load");
            }
        });
    }

    public setLocalStorage(key, value, ttl) {
        var expiryTime = dateAdd(new Date(), "day", ttl);
        const item = {
            value: value,
            expiry: expiryTime,
            currentUser: this.webPartContext.pageContext.user.loginName
        };
        localStorage.setItem(this.webPartContext.pageContext.web.title.toLocaleUpperCase() + "-" + key + "-" + item.currentUser, JSON.stringify(item));
    }

    public getLocalStorage(key) {
        const itemStr = localStorage.getItem(this.webPartContext.pageContext.web.title.toLocaleUpperCase() + "-" + key + "-" + this.webPartContext.pageContext.user.loginName);
        // if the item doesn't exist, return null
        if (!itemStr) {
            return null;
        }
        const item = JSON.parse(itemStr);
        var now = new Date();
        var itemExpiry = new Date(item.expiry);

        // compare the expiry time of the item with the current time
        if (this.webPartContext.pageContext.user.loginName == item.currentUser && now < itemExpiry) {
            return item.value;
        } else {
            localStorage.removeItem(this.webPartContext.pageContext.web.title.toLocaleUpperCase() + "-" + key);
        }
    }


    // to get current user item in UserMaster List
    public getCurrentUserItemID(userTNumber: string): Promise<any> {
        return new Promise<any>(async (resolve: (result: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                this.configValues = await this.getConfigMasterListItems();
                sp.web.lists.getByTitle(UserMasterListName)
                    .items
                    .select("ID,Title,Email")
                    .filter("Title eq '" + userTNumber + "'")
                    .usingCaching({
                        expiration: dateAdd(new Date(), "day", parseInt(this.configValues["UserItemFromUserMasterListCacheExpiry"])),
                        key: this.siteName + "-Home-UserItem-" + this.currentUserName,
                        storeName: "local"
                    })
                    .top(1)
                    .orderBy("ID", false)
                    .get()
                    .then(async (result: any): Promise<void> => {
                        if (result.length == 0) { // to change it to 0
                            localStorage.setItem("isFirstTimeUser", "1");
                            localStorage.removeItem(this.siteName + "-Home-UserItem-" + this.currentUserName);
                            let addedUserItem = await this._addCurrentUserDetailsToList(userTNumber);
                            let addedDefaultMyCollectionsItemIDs = await this._addDefaultMyCollectionsToList(addedUserItem.ID);
                            await this._addUserAndDefaultCollectionToMappingList(addedUserItem.ID, addedDefaultMyCollectionsItemIDs);
                            localStorage.removeItem(this.siteName + "-Home-UserItem-" + this.currentUserName);
                            localStorage.setItem("isFirstTimeUser", "0");
                            resolve(addedUserItem);
                        }
                        else {
                            // update LastAccessedOn for the first time of the day logged in user
                            // if (result[0].LastAccessedOn == null || ConvertToDate(result[0].LastAccessedOn) > 24hrs of currentdatetimenow){
                            // sp.web.lists.getByTitle(UserMasterListName).items.getById(result[0].ID).update({
                            //     LastAccessedOn: new Date()
                            // });
                            // }
                            resolve(result);
                        }
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in getCurrentUserItemID()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getCurrentUserItemID()", error, "Page On Load");
            }
        });
    }

    // to add current user details into UserMaster List for the first time user
    private _addCurrentUserDetailsToList(userTNumber: string): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(UserMasterListName)
                    .items
                    .select("ID")
                    .filter("Title eq '" + userTNumber + "'")
                    .get()
                    .then((result: any): void => {
                        if (result.length > 0) {
                            resolve(result[0]);
                        } else {
                            // adding user details to the UserMaster
                            sp.web.lists.getByTitle(UserMasterListName).items.add({
                                Title: userTNumber.toLocaleLowerCase(),
                                Email: this.webPartContext.pageContext.user.loginName,
                                // LastAccessedOn: new Date()   // add LastAccessedOn for the first time of the day logged in user
                            }).then(async (userAddedResult: IItemAddResult) => {
                                resolve(userAddedResult.data);
                            }, (error: any): void => {
                                reject(error);
                                this.errorLogging.logError(WebPartName, classFileName, "Items in _addCurrentUserDetailsToList()", error, "Page On Load");
                            });
                        }
                    });

            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _addCurrentUserDetailsToList()", error, "Page On Load");
            }
        });
    }

    // to add DefaultMyCollections into UserCollectionMatrix List for the first time user
    private _addDefaultMyCollectionsToList(addedUserItemID: number): Promise<any> {
        let defaultMyCollectionsAddedItemIDs = [];
        return new Promise<any>(async (resolve: (userItem: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                // adding Default My Collection to the CollectionMaster   
                let defaultMyCollections: ICollectionList[] = await this._getDefaultMyCollections();
                let list = sp.web.lists.getByTitle(CollectionMasterListName);
                list.items.select("ID,CollectionOwner/ID")
                    .filter("CollectionOwner/ID eq '" + addedUserItemID + "' and Title eq '" + defaultMyCollections[0].Title + "'")
                    .expand("CollectionOwner")
                    .getAll()
                    .then(async (result: any): Promise<void> => {
                        if (result.length > 0) {
                            defaultMyCollectionsAddedItemIDs.push(result[0].ID);
                            resolve(defaultMyCollectionsAddedItemIDs);
                        }
                        else {
                            // adding multiple items in $batch
                            list.getListItemEntityTypeFullName().then(entityTypeFullName => {
                                let batch = sp.web.createBatch();

                                defaultMyCollections.map((defaultMyCollectionItem: any, index: any) => {
                                    list.items.inBatch(batch).add(
                                        {
                                            Title: defaultMyCollectionItem.Title,
                                            Description: defaultMyCollectionItem.Description,
                                            DefaultMyCollection: 0,
                                            PublicCollection: 0,
                                            CorporateCollection: 0,
                                            StandardOrder: 0,
                                            UnDeletable: 1,
                                            CollectionOwnerId: addedUserItemID

                                        }, entityTypeFullName)
                                        .then((defaultMyCollectionsAddedResult: IItemAddResult) => {
                                            defaultMyCollectionsAddedItemIDs.push(defaultMyCollectionsAddedResult.data.ID);
                                            this.getCurrentCollectionApplications(defaultMyCollections[index]["ID"])
                                                .then((applicationTiles: any) => {
                                                    if (applicationTiles.length > 0) {
                                                        this.addItemToCollectionApplicationMatrixList(applicationTiles, defaultMyCollectionsAddedResult.data.ID);
                                                    }
                                                });
                                        });
                                });

                                batch.execute().then(() => {
                                    resolve(defaultMyCollectionsAddedItemIDs);
                                });

                            }, (error: any): void => {
                                reject(error);
                                this.errorLogging.logError(WebPartName, classFileName, "Items in _addDefaultMyCollectionsToList()", error, "Page On Load");
                            });
                        }
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _addDefaultMyCollectionsToList()", error, "Page On Load");
            }
        });
    }

    // to add CorpCollections too into UserCollectionMatrix List for the first time user
    private _addUserAndDefaultCollectionToMappingList(addedUserItemID: number, addedDefaultMyCollectionsItemIDs: number[]): Promise<any> {

        let success = true;
        return new Promise<any>(async (resolve: (userItem: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                // adding Corporate Collections to the CollectionMaster   
                let corporateCollections: [] = await this._getCorporateCollections();
                corporateCollections = corporateCollections.sort((a: any, b: any) => { return a.StandardOrder - b.StandardOrder; });
                //append Corporate Collection Lists to DefaultMyCollections List
                let defaultMyAndCorporateCollections = addedDefaultMyCollectionsItemIDs.concat(corporateCollections);

                let list = sp.web.lists.getByTitle(UserCollectionMatrixListName);

                list.items.select("UserID/ID,CollectionID/ID")
                    .filter("UserID/ID eq '" + addedUserItemID + "' and CollectionID/ID eq '" + addedDefaultMyCollectionsItemIDs[0] + "'")
                    .expand("UserID,CollectionID")
                    .getAll()
                    .then(async (result: any): Promise<void> => {
                        if (result.length > 0) {
                            resolve(success);
                        }
                        else {
                            // // adding multiple items in $batch
                            list.getListItemEntityTypeFullName().then(entityTypeFullName => {
                                let batch = sp.web.createBatch();

                                defaultMyAndCorporateCollections.map((CollectionItem: any, index: any) => {
                                    list.items.inBatch(batch).add(
                                        {
                                            CollectionOrder: index,
                                            CollectionIDId: CollectionItem.ID != undefined ? CollectionItem.ID : CollectionItem,
                                            UserIDId: addedUserItemID,

                                        }, entityTypeFullName)
                                        .then((addedUserCollectionMatrixItem: IItemAddResult) => {
                                            resolve(addedUserCollectionMatrixItem);
                                        });
                                });
                                batch.execute().then((d) => {
                                    resolve(success);
                                });
                            }, (error: any): void => {
                                reject(error);
                                this.errorLogging.logError(WebPartName, classFileName, "Items in _addUserAndDefaultCollectionToMappingList()", error, "Page On Load");
                            });
                        }
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _addUserAndDefaultCollectionToMappingList()", error, "Page On Load");
            }
        });
    }

    //to get all current user collection ids 
    private _getCurrentUserCollectionIDs(userItem: any): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                let userItemID = userItem.ID != undefined ? userItem.ID : userItem[0].ID;
                sp.web.lists.getByTitle(UserCollectionMatrixListName)
                    .items
                    .orderBy("ID")
                    .select("ID,CollectionID/ID,CollectionOrder")
                    .expand("CollectionID")
                    .filter("UserID/ID eq '" + userItemID + "'")
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in _getCurrentUserCollectionIDs()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _getCurrentUserCollectionIDs()", error, "Page On Load");
            }
        });
    }

    //to get DefaultMyCollections from CollectionMaster List
    private _getDefaultMyCollections(): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionMasterListName)
                    .items
                    .orderBy("StandardOrder")
                    .select("ID,Title,Description,StandardOrder")
                    .filter("(DefaultMyCollection eq 1) and (CorporateCollection ne 1)")
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in _getDefaultMyCollections()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _getDefaultMyCollections()", error, "Page On Load");
            }
        });
    }

    //to get CorporateCollections from CollectionMaster List
    private _getCorporateCollections(): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionMasterListName)
                    .items
                    .orderBy("StandardOrder")
                    .select("ID,Title,Description,StandardOrder")
                    .filter("(CorporateCollection eq 1)")
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in _getCorporateCollections()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _getCorporateCollections()", error, "Page On Load");
            }
        });
    }

    //to get all the applications from CollectionAppMatrixList based on Current Collection
    public getCurrentCollectionApplications(collectionItemID: any): Promise<any> {
        var applicationTileItems: IApplicationList[] = [];
        return new Promise<any>((resolve: (applicationTileItems: IApplicationList[]) => void, reject: (error: any) => void): void => {
            try {
                this._getApplicationIDsBasedOnCurrentCollection(collectionItemID)
                    .then((collectionApplicationItems: any) => {
                        if (collectionApplicationItems.length > 0) {
                            sp.web.lists.getByTitle(ApplicationMasterListName)
                                .items
                                .orderBy("Title")
                                .select("ID,Title,Description,OwnerEmail,Link,SearchKeywords,AvailableExternal,ColorCode/Title,ColorCode/BgColor,ColorCode/ForeColor,IsActive")
                                .expand("ColorCode")
                                //.filter(this._generateFilterCondition("ID", "ApplicationID", collectionApplicationItems) + " and (IsActive eq 1)")
                                .filter("IsActive eq 1")
                                .getAll()
                                .then((result: any): void => {
                                    //Filtering all application master items based on its it presented in collectionApplicationItems application id
                                    result = result.filter(a => true === collectionApplicationItems.some(b => a.ID === b.ApplicationID.ID));

                                    result.map((value: any) => {
                                        var matchedApplicationCollectionItem = collectionApplicationItems.filter((collectionApplicationItem: any) => {
                                            return (collectionApplicationItem["ApplicationID"]["ID"] == value["ID"]);
                                        });
                                        if (matchedApplicationCollectionItem.length > 0) {
                                            value["AppOrder"] = matchedApplicationCollectionItem[0]["AppOrder"];
                                            value["ApplicationCollectionMatrixID"] = matchedApplicationCollectionItem[0]["ID"];
                                            applicationTileItems.push(value);
                                        }
                                        else {
                                            value["AppOrder"] = 0;
                                            applicationTileItems.push(value);
                                        }
                                    });
                                    resolve(applicationTileItems);
                                }, (error: any): void => {
                                    reject(error);
                                    this.errorLogging.logError(WebPartName, classFileName, "Items in getCurrentCollectionApplications()", error, "Page On Load");
                                });
                        } else {
                            resolve(applicationTileItems);
                        }
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getCurrentCollectionApplications()", error, "Page On Load");
            }
        });
    }

    //to get further application details from ApplicationMaster based on ApplicationIDs
    private _getApplicationIDsBasedOnCurrentCollection(collectionItemID: any): Promise<any> {
        return new Promise<any>((resolve: (applicationTileItems: ICollectionList) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionApplicationMatrixListName)
                    .items
                    .orderBy("ID")
                    .select("ID,ApplicationID/ID,AppOrder")
                    .expand("ApplicationID")
                    .filter("CollectionID/ID eq '" + collectionItemID + "'")
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in _getApplicationIDsBasedOnCurrentCollection()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in _getApplicationIDsBasedOnCurrentCollection()", error, "Page On Load");
            }
        });
    }

    //to add PrivateCollection to CollectionMAster List
    public addPrivateCollectionsToMasterList(collectionItem: any): Promise<any> {
        return new Promise<any>((resolve: (collectionItem: ICollectionList) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionMasterListName).items
                    .add(collectionItem)
                    .then(async (collectionItemAddedResult: IItemAddResult) => {
                        resolve(collectionItemAddedResult.data);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in addPrivateCollectionsToMasterList()", error, "Create a new Collection");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in addPrivateCollectionsToMasterList()", error, "Create a new Collection");
            }
        });
    }

    //to add mapping items to the UserCollectionMatrix List
    public addItemToUserCollectionMatrixList(userCollectionItem: any): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(UserCollectionMatrixListName).items
                    .add(userCollectionItem)
                    .then(async (userCollectionItemResult: IItemAddResult) => {
                        resolve(userCollectionItemResult.data);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in addItemToUserCollectionMatrixList()", error, "Create a new Collection");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in addItemToUserCollectionMatrixList()", error, "Create a new Collection");
            }
        });
    }

    //to add mapping items to the CollectionAppMatrix List
    public addItemToCollectionApplicationMatrixList(applicationTileItems: any, collectionListItemID: any, currentMaximumApplicationOrder?: number): Promise<any> {
        let collectionApplicationMatrixItemIDs = [];
        return new Promise<any>((resolve: (collectionApplicationMatrixListName: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(CollectionApplicationMatrixListName);
                // adding multiple items in $batch
                list.getListItemEntityTypeFullName().then(entityTypeFullName => {
                    let batch = sp.web.createBatch();

                    applicationTileItems.map((applicationTileItem: any, index: any) => {
                        list.items.inBatch(batch).add(
                            {
                                AppOrder: currentMaximumApplicationOrder != undefined ? currentMaximumApplicationOrder + index + 1 : index,
                                ApplicationIDId: applicationTileItem.ID,
                                CollectionIDId: collectionListItemID
                            }, entityTypeFullName)
                            .then((applicationTileItemAddedResult: IItemAddResult) => {
                                collectionApplicationMatrixItemIDs.push(applicationTileItemAddedResult.data.ID);
                            });
                    });
                    batch.execute().then(() => {
                        resolve(collectionApplicationMatrixItemIDs);
                    });
                }, (error: any): void => {
                    reject(error);
                    this.errorLogging.logError(WebPartName, classFileName, "Items in addItemToCollectionApplicationMatrixList()", error, "Create a new Collection");
                });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in addItemToCollectionApplicationMatrixList()", error, "Create a new Collection");
            }
        });
    }

    //to add Public Collection to Request List
    public addPublicCollectionsToRequestList(collectionRequestItem: any): Promise<any> {
        return new Promise<any>((resolve: (userItem: ICollectionList) => void, reject: (error: any) => void): void => {
            try {
                collectionRequestItem.RequestedDate = new Date();
                sp.web.lists.getByTitle(CollectionRequestsListName).items
                    .add(collectionRequestItem)
                    .then(async (collectionItemAddedResult: IItemAddResult) => {
                        resolve(collectionItemAddedResult.data);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in addPublicCollectionsToRequestList()", error, "Create a new Collection");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in addPublicCollectionsToRequestList()", error, "Create a new Collection");
            }
        });
    }

    //to add Public Collection's application to CollectionApplicationMatrixRequest List
    public addItemToCollectionApplicationMatrixRequestsList(applicationTileItems: any, collectionRequestItemID: any): Promise<any> {
        let collectionApplicationMatrixRequestsItemIDs = [];
        return new Promise<any>((resolve: (collectionApplicationMatrixListName: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(CollectionApplicationMatrixRequestsListName);
                // adding multiple items in $batch
                list.getListItemEntityTypeFullName().then(entityTypeFullName => {
                    let batch = sp.web.createBatch();

                    applicationTileItems.map((applicationTileItem: any, index: any) => {
                        list.items.inBatch(batch).add(
                            {
                                ApplicationIDId: applicationTileItem.ID,
                                CollectionRequestIDId: collectionRequestItemID
                            }, entityTypeFullName)
                            .then((applicationTileItemAddedResult: IItemAddResult) => {
                                collectionApplicationMatrixRequestsItemIDs.push(applicationTileItemAddedResult.data.ID);
                            });
                    });
                    batch.execute().then(() => {
                        resolve(collectionApplicationMatrixRequestsItemIDs);
                    });
                }, (error: any): void => {
                    reject(error);
                    this.errorLogging.logError(WebPartName, classFileName, "Items in addItemToCollectionApplicationMatrixRequestsList()", error, "Create a new Collection");
                });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in addItemToCollectionApplicationMatrixRequestsList()", error, "Create a new Collection");
            }
        });
    }

    public addMulipleItemsToUserCollectionMatrixList(userItemId: any, collectionListItems: any, currentMaximumCollectionOrder: number): Promise<any> {
        let userCollectionMatrixListMatrixItemIDs = [];
        return new Promise<any>((resolve: (UserCollectionMatrixListName: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(UserCollectionMatrixListName);
                // adding multiple items in $batch
                list.getListItemEntityTypeFullName().then(entityTypeFullName => {
                    let batch = sp.web.createBatch();

                    collectionListItems.map((collectionListItem: any, index: any) => {
                        list.items.inBatch(batch).add(
                            {
                                CollectionIDId: collectionListItem.ID,
                                UserIDId: userItemId[0].ID,
                                CollectionOrder: currentMaximumCollectionOrder + index + 1
                            }, entityTypeFullName)
                            .then((collectionListItemsAddedResult: IItemAddResult) => {
                                userCollectionMatrixListMatrixItemIDs.push(collectionListItemsAddedResult.data.ID);
                            });
                    });
                    batch.execute().then(() => {
                        resolve(userCollectionMatrixListMatrixItemIDs);
                    });
                }, (error: any): void => {
                    reject(error);
                    this.errorLogging.logError(WebPartName, classFileName, "Items in addMulipleItemsToUserCollectionMatrixList()", error, "Create a new Collection");
                });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in addMulipleItemsToUserCollectionMatrixList()", error, "Create a new Collection");
            }
        });
    }

    //to delete items on UserCollection Matrix List
    public deleteItemsOnUserCollectionMatrixList(userCollectionItems: any): Promise<any> {
        let success: true;
        let deleteCount: number = 0;
        return new Promise<any>((resolve: (UserCollectionMatrix: any) => void, reject: (error: any) => void): void => {
            try {
                if (userCollectionItems.length > 0) {
                    let list = sp.web.lists.getByTitle(UserCollectionMatrixListName);
                    userCollectionItems.map((userCollectionItem: any) => {
                        list.items.getById(userCollectionItem.UserCollectionMatrixItemID)
                            .delete()
                            .then(() => {
                                deleteCount++;
                                if (userCollectionItems.length == deleteCount)
                                    resolve(success);
                            });
                    });
                }
                else {
                    resolve(success);
                }
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Items in deleteItemsOnUserCollectionMatrixList()", error, "Delete a Collection");
            }
        });
    }

    //to update Sorting Order in UserCollectionMatrix List
    public updateSortingOrderInUserCollectionMatrixList(userCollectionItems: any): Promise<any> {
        let success: true;
        return new Promise<any>((resolve: (UserCollectionMatrixItem: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(UserCollectionMatrixListName);
                let batch = sp.web.createBatch();
                userCollectionItems.map((userCollectionItem: any, index: any) => {
                    list.items.getById(userCollectionItem.UserCollectionMatrixItemID)
                        .inBatch(batch)
                        .update({
                            CollectionOrder: index
                        });
                });
                batch.execute().then(() => {
                    resolve(success);
                });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Items in updateSortingOrderInUserCollectionMatrixList()", error, "Left Navigation Collection ReOrdering");
            }
        });
    }

    //to update Sorting Order in CollectionApplicationMatrix List
    public updateSortingOrderInCollectionApplicationMatrixList(applicationItems: any): Promise<any> {
        let success: true;
        return new Promise<any>((resolve: (CollectionApplicationMatrixItem: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(CollectionApplicationMatrixListName);
                let batch = sp.web.createBatch();
                applicationItems.map((applicationItem: any, index: any) => {
                    list.items.getById(applicationItem.ApplicationCollectionMatrixID)
                        .inBatch(batch)
                        .update({
                            AppOrder: index
                        });
                });
                batch.execute().then(() => {
                    resolve(success);
                });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Items in updateSortingOrderInCollectionApplicationMatrixList()", error, "Tile ReOrdering in a Collection");
            }
        });
    }

    //to get current collection details
    public getCurrentCollectionDetails(currentCollectionID: string, myFollowedCollections: any): Promise<any> {
        let currentCollectionItem: ICollectionList[] = [];
        return new Promise<any>((resolve: (isValidCollectionID: any) => void, reject: (error: any) => void): void => {
            try {
                currentCollectionItem = myFollowedCollections.filter((collection: any) => String(collection.ID) === currentCollectionID);
                if (currentCollectionItem.length > 0) {
                    resolve(currentCollectionItem[0]);
                }
                else {
                    sp.web.lists.getByTitle(CollectionMasterListName)
                        .items
                        .select("ID,Title,Description,PublicCollection,DefaultMyCollection,CorporateCollection,StandardOrder,UnDeletable,CollectionOwner/Title,IsActive")
                        .expand("CollectionOwner")
                        .filter("ID eq " + currentCollectionID)
                        .getAll()
                        .then((result: any): void => {
                            if (result.length == 1) {
                                result[0]["CollectionOwner"] = result[0]["CollectionOwner"]["Title"];
                            }
                            resolve(result[0]);
                        }, (error: any): void => {
                            reject(error);
                            this.errorLogging.logError(WebPartName, classFileName, "Items in getCurrentCollectionDetails()", error, "On Callback Page Re-rendering");
                        });
                }
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getCurrentCollectionDetails()", error, "On Callback Page Re-rendering");
            }
        });
    }

    //to get configmasterList items
    public getConfigMasterListItems(): Promise<any> {
        let configMasterItems = new Array();
        return new Promise<any>((resolve: (item: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(ConfigMasterListName)
                    .items
                    .select("Title", "ConfigValue")
                    .top(5000)
                    .usingCaching({
                        expiration: dateAdd(new Date(), "day", 7),
                        key: this.siteName + "-ConfigMaster",
                        storeName: "local"
                    })
                    .get()
                    .then((items: any) => {
                        try {
                            if (items.length > 0) {
                                items.map((value: any, index: any) => {
                                    configMasterItems[value["Title"]] = value["ConfigValue"];
                                });
                                resolve(configMasterItems);
                            }
                        }
                        catch (error) {
                            reject(error);
                        }
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "Items in getConfigMasterListItems()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getConfigMasterListItems()", error, "Page On Load");
            }
        });
    }

    //to check current logged in user is part of PGOneApprover group
    public checkCurrentUserApprovalPermission(): Promise<any> {
        let isCurrentUserApprover: boolean = false;
        return new Promise<any>(async (resolve: (item: any) => void, reject: (error: any) => void): Promise<void> => {
            try {
                sp.web.currentUser
                    .groups()
                    .then((result: any): void => {
                        if (result.length > 0) {
                            isCurrentUserApprover = result.some((group: any) => {
                                return group.LoginName === this.configValues["PGOneApproverGroup"];
                            });
                        }
                        resolve(isCurrentUserApprover);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "checkCurrentUserApprovalPermission()", error, "Page On Load");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in checkCurrentUserApprovalPermission()", error, "Page On Load");
            }
        });
    }

    //to delete items in CollectionAppMatrixList
    public deleteItemsOnCollectionApplicationMatrixList(applicationItems: any): Promise<any> {
        let success: true;
        let deleteCount: number = 0;
        return new Promise<any>((resolve: (CollectionApplicationMatrixItem: any) => void, reject: (error: any) => void): void => {
            try {
                if (applicationItems.length > 0) {
                    let list = sp.web.lists.getByTitle(CollectionApplicationMatrixListName);
                    applicationItems.map((applicationItem: any) => {
                        list.items.getById(applicationItem.ApplicationCollectionMatrixID)
                            .delete()
                            .then(() => {
                                deleteCount++;
                                if (applicationItems.length == deleteCount)
                                    resolve(success);
                            });
                    });
                }
                else {
                    resolve(success);
                }
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in deleteItemsOnCollectionApplicationMatrixList()", error, "Delete a Collection");
            }
        });
    }

    // to update record in CollectionMaster List
    public updateCollectionSettings(collectionItem: any): Promise<any> {
        let success: true;
        return new Promise<any>((resolve: (CollectionItem: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(CollectionMasterListName);
                list.items.getById(collectionItem.ID)
                    .update({
                        Title: collectionItem.Title,
                        Description: collectionItem.Description
                    }).then(() => {
                        resolve(success);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "updateCollectionSettings()", error, "Update a Collection Settings");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in updateCollectionSettings()", error, "Update a Collection Settings");
            }
        });
    }

    // to get public Collection request details based on collecitonID and disable the panel if already awaiting approval
    public getPublicCollectionRequestDetailsBasedOnCollectionID(collectionItem: any): Promise<any> {
        let collectionRequestAvailable: boolean = false;
        return new Promise<any>((resolve: (collectionRequestAvailable: boolean) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionRequestsListName)
                    .items
                    .orderBy("ID")
                    .select("ID")
                    .filter("ExistingItemID eq '" + collectionItem.ID + "' and ApprovalStatus eq 'Waiting For Approval'")
                    .getAll()
                    .then((result: any): void => {
                        if (result.length > 0) {
                            collectionRequestAvailable = true;
                            resolve(collectionRequestAvailable);
                        }
                        else {
                            resolve(collectionRequestAvailable);
                        }
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "getPublicCollectionRequestDetailsBasedOnCollectionID()", error, "Update a Collection Settings");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getPublicCollectionRequestDetailsBasedOnCollectionID()", error, "Update a Collection Settings");
            }
        });
    }

    //to delete collection item on collection master list
    public deleteCollectionItemOnCollectionMasterList(collectionItem: any): Promise<any> {
        let success: true;
        return new Promise<any>((resolve: (success: any) => void, reject: (error: any) => void): void => {
            try {
                let list = sp.web.lists.getByTitle(CollectionMasterListName);
                list.items
                    .getById(collectionItem.ID)
                    .update({
                        IsActive: false
                    })
                    .then(() => {
                        resolve(success);
                    }, (error: any): void => {
                        reject(error);
                        this.errorLogging.logError(WebPartName, classFileName, "deleteCollectionItemOnCollectionMasterList()", error, "Delete a Collection");
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in deleteCollectionItemOnCollectionMasterList()", error, "Delete a Collection");
            }
        });
    }

    // Savan
    //#region  Manage My Tiles

    public getMyTiles(filterCondition: string): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(ApplicationMasterListName)
                    .items


                    .select("ID,Title,Description,SearchKeywords,AvailableExternal,OwnerEmail,Link,IsActive,ApprovedDate,ColorCode/Title,ColorCode/BgColor,ColorCode/ForeColor,ColorCode/ID,ColorCode/Title")
                    .expand('ColorCode')
                    .filter(filterCondition)
                    .orderBy("Modified", false)
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getMyTiles()", error, "Get Tiles");
            }
        });
    }


    public createTileRequestToApplicationRequest(items: any): Promise<any> {
        return new Promise<any>(async (resolve: (IApplicationList: any) => void, reject: (error: any) => void): Promise<void> => {
            try {

                // adding user details to the UserMaster
                sp.web.lists.getByTitle(ApplicationRequestsListName).items.add({
                    Title: items['TileTitle'],
                    OwnerEmail: items['TileOwner'],
                    Link: items['TileURLLink'],
                    Description: items['TileDescription'],
                    SearchKeywords: items['TileKewords'],
                    ExistingItemID: parseInt(items['TileId']),
                    ColorCodeId: parseInt(items['TileColorCodeId']),
                    RequestedDate: new Date(),
                    RequestedAction: items['RequestedAction'],
                    AvailableExternal: items['TileAvailableExternal']
                }).then((result: any): void => {
                    resolve(result);
                }, (error: any): void => {
                    reject(error);
                });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in createTileRequestToApplicationRequest()", error, "Create Collection Request");
            }
        });
    }

    public checkExitingRecordIsInApprovalInTrackRequest = (itemId: string): Promise<any> => {
        let itemFound = false;
        return new Promise<any>((resolve: (item: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(ApplicationRequestsListName)
                    .items
                    .filter("ExistingItemID eq " + itemId + " and ApprovalStatus eq 'Waiting for Approval'")
                    .get()
                    .then((result) => {
                        try {
                            itemFound = result.length > 0 ? true : false;
                            resolve(itemFound);
                        }
                        catch (error) {
                            reject(error);
                        }

                    }, (error: any): void => {
                        console.log(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in checkExitingRecordIsInApprovalInTrackRequest()", error, "Check Exiting RecordIs In Approval In TrackRequest");
            }
        });

    }


    public checkApplicationUrlExistOrNot = (itemId: string, url: string): Promise<any> => {
        let urlFound = false;
        return new Promise<any>((resolve: (item: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(ApplicationMasterListName)
                    .items.top(5000)
                    .filter("Id ne " + itemId)
                    .get()
                    .then((result) => {
                        try {
                            urlFound = result.some(tile => { return tile.Link.toLowerCase() === url.toLowerCase(); });
                            resolve(urlFound);
                        }
                        catch (error) {
                            reject(error);
                        }

                    }, (error: any): void => {
                        console.log(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in checkApplicationUrlExistOrNot()", error, "Check Application Url ExistOrNot");
            }
        });
    }

    //#endregion Manage my tiles

    //#region Track Request
    public getTrackRequestForApplication(filterCondition: string): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(ApplicationRequestsListName)
                    .items
                    .select("ID,Title,Description,AvailableExternal,SearchKeywords,OwnerEmail,Link,ColorCode/BgColor,ColorCode/ForeColor,ColorCode/ID,ColorCode/Title,ApprovalStatus,RequestedAction,RequestedDate,DecisionDate,DecisionBy,DecisionComments")
                    .expand('ColorCode')
                    .filter(filterCondition)
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getTrackRequestForApplication()", error, "Get TrackRequest For Application");
            }
        });
    }

    public getTrackRequestForCollection(filterCondition: string): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionRequestsListName)
                    .items
                    .select("ID,Title,Description,PublicCollection,RequestedDate,ApprovalStatus,RequestedAction,CollectionOwner/Email,CollectionOwner/Id,ExistingItemID,DecisionDate,DecisionBy,DecisionComments")
                    .expand('CollectionOwner')
                    .filter(filterCondition)
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getTrackRequestForCollection()", error, "Get TrackRequest For collection");
            }
        });
    }

    public getCurrentCollectionApplicationFromCollectionApplicationMatrx(ExID: number): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionApplicationMatrixListName)
                    .items
                    .select("ID,Title,CollectionIDId,ApplicationIDId")
                    .filter("CollectionIDId eq " + ExID + "")
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getCurrentCollectionApplicationFromCollectionApplicationMatrx()", error, "Get Current Collection Application From Collection pplicationMatrx");
            }
        });
    }

    public getCurrentCollectionApplication(ID: number): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(CollectionApplicationMatrixRequestsListName)
                    .items
                    .select("ID,Title,CollectionRequestIDId,ApplicationIDId")
                    .filter("CollectionRequestIDId eq " + ID + "")
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getCurrentCollectionApplication()", error, "Get Current Collection Application");
            }
        });
    }
    public getMyTilesByItsId(filterCondition: string): Promise<any> {
        return new Promise<any>((resolve: (userItem: any) => void, reject: (error: any) => void): void => {
            try {
                sp.web.lists.getByTitle(ApplicationMasterListName)
                    .items
                    .select("ID,Title,Description,SearchKeywords,AvailableExternal,OwnerEmail,Link,IsActive,ApprovedDate,ColorCode/Title,ColorCode/BgColor,ColorCode/ForeColor,ColorCode/ID")
                    .expand('ColorCode')
                    .filter(filterCondition)
                    .getAll()
                    .then((result: any): void => {
                        resolve(result);
                    }, (error: any): void => {
                        reject(error);
                    });
            }
            catch (error) {
                reject(error);
                this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in getMyTilesByItsId()", error, "Get MyTiles By Its Id");
            }
        });
    }
    //#endregion Track Request
    // end Savan

    /*Added By Sunny */
    public createTileRequest(itemTileRequest: ITileRequest): Promise<any> {
        return new Promise<any>(
            async (
                resolve: (item: any) => void,
                reject: (error: any) => void
            ): Promise<void> => {
                try {
                    sp.web.lists
                        .getByTitle(ApplicationRequestsListName)
                        .items.add(itemTileRequest)
                        .then(
                            (r: any): void => {
                                console.log(r);
                                resolve(r);
                            },
                            (error: any): void => {
                                reject(error);
                            }
                        );
                } catch (error) {
                    reject(error);
                    this.errorLogging.logError(WebPartName, classFileName, "Some Error Occured in createTileRequest()", error, "Create a new tile request");
                }
            }
        );
    }

    public reSetDefaultProfile(tNumber: string): Promise<any> {
        return new Promise<any>(async (resolve: (userItem: any) => void, reject: (error: any) => void): Promise<void> => {
                //remove items from collection master
                try {
                    var userId = await this.getCurrentUserItemID(tNumber);
                    let listColl = sp.web.lists.getByTitle(CollectionMasterListName);
                    let listUserMaster = sp.web.lists.getByTitle(UserMasterListName);
                    listColl.items
                        .orderBy("ID")
                        .select("ID,Title,CollectionOwner/ID,CollectionOwner/Title")
                        .expand("CollectionOwner")
                        .filter("CollectionOwner/ID eq '" + userId[0].ID + "'")
                        .getAll()
                        .then((result: any): void => {
                            //resolve(result);
                            console.log(result);
                            //Delete All items of collection master
                            if (result.length > 0) 
                            {
                                result.map((Items) => {
                                    listColl.items.getById(Items.Id)
                                        .delete()
                                        .then(() => {
                                            //console.log(Items.Id, "collection deleted");
                                        });
                                });
                            }
                            //remove items from user master
                            listUserMaster.items.getById(userId[0].ID)
                            .delete()
                            .then(() => {
                                //console.log(userId[0].ID , "User master deleted");
                                resolve("Profile Cleared");
                            });
                            
                        }, (error: any): void => {
                            reject(error);
                        });
                }
                catch (error) {
                    reject(error);
                }

        });
    }
}
