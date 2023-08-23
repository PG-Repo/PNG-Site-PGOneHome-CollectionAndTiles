import { initialize, pageview, ga, OutboundLink } from 'react-ga';


export class gAnalytics {

    constructor(trackingId: string, userId: string) {
        if (trackingId.trim() !== "") {
            initialize(trackingId, {
                debug: false,
                titleCase: false,
                gaOptions: {
                    clientId: userId,
                    userId: userId,
                    siteSpeedSampleRate: 100,
                }
            });
            pageview(window.location.pathname + window.location.search);
        }
    }

    //Google Analytics: collection onclick
    public collectionHit = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection click from user list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }

    //Google Analytics: Collection Hit from Querystring
    public collectionHitFromGet = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection hit from direct link', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }

    //Google Analytics: Collection Hit from List (Edit which collection you follow) 
    public collectionHitFromList = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection hit from list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }

    //Google Analytics: Collection followed with Direct link 
    public collectionDirectFollow = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection direct follow', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }

    //Google Analytics: Collection unfollowed with Direct link
    public collectionDirectUnfollow = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection direct unfollow', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }

    //Google Analytics: collection checked from list (Edit which collections you follow)
    public collectionListFollow = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection follow from list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }
    //Google Analytics: collection unchecked from list (Edit which collections you follow)
    public collectionListUnfollow = (name, id) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection unfollow from list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': name,
            'dimension9': id
        });
    }
    //Google Analytics: tile Hit in collection (ApplicationTiles)
    public appHitInCollection = (name, link, collection, collectionId, appId, isAppRemoveClicked: boolean) => {
        //if (!$(event.target).is('.removeAppButton')) {
        if (!isAppRemoveClicked) {
            try {
                ga('send', {
                    'hitType': 'event', // Required.
                    'eventCategory': 'Applications', // Required.
                    'eventAction': 'App click in collection', // Required.
                    'eventLabel': name,
                    'eventValue': 1,
                    'dimension8': collection,
                    'dimension9': collectionId,
                    'dimension10': appId,
                    'dimension13': link
                });
            } catch (err) {
                //don't care, here so we can continue in case of failure
            }
        }
    }
    //Google Analytics: tile link visited from list (Create a new collection)
    public appHitFromList = (link, appId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Applications', // Required.
            'eventAction': 'App click in list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension10': appId,
            'dimension13': link
        });

    }

    //Google Analytics: remove tile from list (Application Tiles)
    public appDirectRemove = (name, link, collection, collectionId, appId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'App removed directly from collection', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId,
            'dimension10': appId,
            'dimension13': link
        });
    }

    //Google Analytics: tile checked from list (Create a new collection)
    public appListAdd = (name, link, collection, collectionId, appId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'App box checked from list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId,
            'dimension10': appId,
            'dimension13': link
        });
    }
    //Google Analytics: tile unchecked from list (Create a new collection)
    public appListRemove = (name, link, collection, collectionId, appId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'App box unchecked from list', // Required.
            'eventLabel': name,
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId,
            'dimension10': appId,
            'dimension13': link
        });
    }

    //Google Analytics: tile search from list (Create a new collection)
    public appListSearch = (collection, collectionId, searchTerm) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Applications', // Required.
            'eventAction': 'App list search', // Required.
            'eventLabel': 'searchTerm',
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId,
            'dimension12': searchTerm
        });
    }

    //Google Analytics: collection searched from list (Edit which collections you follow)
    public collectionListSearch = (searchTerm) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection list search', // Required.
            'eventLabel': 'searchTerm',
            'eventValue': 1,
            'dimension12': searchTerm
        });
    }

    //Google Analytics: share collection clicked
    public shareCollectionClick = (collection, collectionId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Share collection click', // Required.
            'eventLabel': collection,
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId
        });
    }

    //Google Analytics:collection reorder fired
    public collectionReorderFired = () => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collections reordered', // Required.
            'eventLabel': 'Collections reordered',
            'eventValue': 1
        });
    }

    //Google Analytics:tile reorder fired
    public appReorderFired = (collection, collectionId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Applications reordered', // Required.
            'eventLabel': 'Applications reordered',
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId
        });
    }

    //Google Analytics: create Collection Called
    public createCollectionCalled = () => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Create collection button hit', // Required.
            'eventLabel': 'Create collection button hit',
            'eventValue': 1
        });
    }

    //Google Analytics: create Collection Save
    public createCollectionSave = () => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Create collection save', // Required.
            'eventLabel': 'Create collection save',
            'eventValue': 1
        });
    }

    //Google Analytics: collection Settings Called
    public collectionSettingsCalled = (collection, collectionId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection settings button clicked', // Required.
            'eventLabel': collection,
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId
        });
    }

    //Google Analytics: collection Settings Save Called
    public collectionSettingsSave = (collection, collectionId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Collection settings save', // Required.
            'eventLabel': collection,
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId
        });
    }

    //Google Analytics: Edit which collections you follow
    public editCollectionsFollowedClicked = () => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Edit collections followed clicked', // Required.
            'eventLabel': 'Click',
            'eventValue': 1
        });
    }

    //Google Analytics: Edit which collections you follow save clicked
    public editCollectionsFollowedSaved = () => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Edit collections followed saved', // Required.
            'eventLabel': 'Edit collections followed saved',
            'eventValue': 1
        });
    }

    //Google Analytics: edit Tiles In Collection Clicked
    public editSitesInCollectionClicked = (collection, collectionId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Edit sites in collection clicked', // Required.
            'eventLabel': 'Edit sites in collection clicked',
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId
        });
    }

    //Google Analytics: edit Tiles In Collection Save Clicked
    public editSitesInCollectionSaved = (collection, collectionId) => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Collections', // Required.
            'eventAction': 'Edit sites in collection saved', // Required.
            'eventLabel': 'Edit sites in collection saved',
            'eventValue': 1,
            'dimension8': collection,
            'dimension9': collectionId
        });
    }

    //Google Analytics: Try Link Clicked
    public appTryLinkClicked = () => {
        ga('send', {
            'hitType': 'event', // Required.
            'eventCategory': 'Applications', // Required.
            'eventAction': 'Try anyway link clicked', // Required.
            'eventLabel': 'Try anyway link clicked',
            'eventValue': 1,
        });
    }

}
