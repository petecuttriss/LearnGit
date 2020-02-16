import {
    ClientSideText,
    ClientSideWebpart,
    sp,
    Web,
    ClientSidePage,
    SPRest
} from "@pnp/sp";
import { PageContext } from "@microsoft/sp-page-context";
import { Environment, EnvironmentType, ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { stringIsNullOrEmpty } from "@pnp/common";
import { IListService } from "../models/IListService";
import { Utils } from "./Utils";

export class ListService implements IListService {

    public static readonly serviceKey: ServiceKey<IListService> = ServiceKey.create<IListService>('spfx-teams-web-parts:ListService', ListService);

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {

            const pageContext = serviceScope.consume(PageContext.serviceKey);

            // we need to "spoof" the context object with the parts we need for PnPjs
            sp.setup({
                spfxContext: {
                    pageContext: pageContext,
                }
            });
        });
    }

    public async getPropertyBagItem(item: string): Promise<string> {
        var propertyBagItem: string = "Water Supply";

        if (Environment.type === EnvironmentType.Local) {
            propertyBagItem = this.getPropertyBagItemFromMock();
        } else {
            propertyBagItem = await this.getPropertyBagItemFromSP(item);
        }

        return propertyBagItem;
    }

    private getPropertyBagItemFromMock() {
        return "Water Supply";
    }

    private async getPropertyBagItemFromSP(item: string) {
        var propertyBagItem: string = "";
        var propertyBagItemKey: string = this.getEncodedString(item);

        var propertyBagItems: any = await sp.web.select("AllProperties")
            .expand("AllProperties")
            .get();

        if (propertyBagItems.AllProperties[propertyBagItemKey] != null) {
            propertyBagItem = propertyBagItems.AllProperties[propertyBagItemKey];
        }

        return propertyBagItem;
    }

    private getEncodedString(item: string): string {
        var encodedString = "";
        if (!stringIsNullOrEmpty(item) && item !== undefined) {
            var charToEncode = item.split('');

            for (let i = 0; i < charToEncode.length; i++) {
                let encodedChar = escape(charToEncode[i]);

                if (encodedChar.length == 3) {
                    encodedString += encodedChar.replace("%", "_x00") + "_";
                }
                else if (encodedChar.length == 5) {
                    encodedString += encodedChar.replace("%u", "_x") + "_";
                }
                else if (encodedChar == ".") {
                    encodedString += ("_x002e_");
                }
                else {
                    encodedString += encodedChar;
                }
            }
        }
        return encodedString;
    }

    public getItems(siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any> {
        if (Environment.type === EnvironmentType.Local) {
            return this.getItemsFromMock(groupBy);
        } else {
            return this.getItemsFromSP(siteUrl, listId, fields, groupBy, filter);
        }
    }

    private getItemsFromSP(siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any> {
        let web = new Web(siteUrl);
        let filterValue = filter === undefined ? "" : filter;
        return web.lists.getById(listId).items.select(fields).orderBy("Title", true).filter(filterValue).get().then((items) => {
            let groupedItems = Utils.groupBy(items, item => item[groupBy]);
            let sortedItems = Utils.sortByKey(groupedItems);
            return sortedItems;
        }).catch((error: any) => {
            let errorMessage: string = "";
            if (error.status === 403) {
                // Still resolve with accessDenied=true
                errorMessage = "Access Denied";
            }

            // If it fails because previously configured web/list doesn't exist anymore
            else if (error.status === 404) {
                // Still resolve with site not found
                errorMessage = "Site Not Found";
            }

            // If it fails for any other reason, reject with the error message
            else {
                errorMessage = error.statusText ? error.statusText : error.message;
            }
            return { "error": errorMessage };
        });
    }

    private getItemsFromMock(groupBy: string): Promise<any> {
        let mockItems: any = require("./items.json");
        let groupedItems = Utils.groupBy(mockItems, item => item[groupBy]);
        let sortedItems = Utils.sortByKey(groupedItems);

        return new Promise<any>((resolve) => {
            //setTimeout(() => {
            resolve(sortedItems);
            //}, 2000);
        });
    }

    public searchItems(query: string, siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any> {
        if (Environment.type === EnvironmentType.Local) {
            return this.searchItemsFromMock(query, groupBy);
        } else {
            return this.searchItemsFromSP(query, siteUrl, listId, fields, groupBy, filter);
        }
    }

    private searchItemsFromSP(query: string, siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any> {
        let web = new Web(siteUrl);
        let filterValue = filter === undefined ? "" : filter;
        return web.lists.getById(listId).items.select(fields).orderBy("Title", true).filter(filterValue).get().then((items) => {
            let searchResults = items.filter((item) => {
                return item.Title.toLowerCase().search(
                    query.toLowerCase()) !== -1; //|| item.SFDesc.toLowerCase().search(query.toLowerCase()) !== -1;
            });
            let groupedItems = Utils.groupBy(searchResults, item => item[groupBy]);
            let sortedItems = Utils.sortByKey(groupedItems);
            return sortedItems;
        }).catch((error: any) => {
            let errorMessage: string = "";
            if (error.status === 403) {
                // Still resolve with accessDenied=true
                errorMessage = "Access Denied";
            }

            // If it fails because previously configured web/list doesn't exist anymore
            else if (error.status === 404) {
                // Still resolve with site not found
                errorMessage = "Site Not Found";
            }

            // If it fails for any other reason, reject with the error message
            else {
                errorMessage = error.statusText ? error.statusText : error.message;
            }
            return { "error": errorMessage };
        });
    }

    private searchItemsFromMock(query: string, groupBy: string): Promise<any> {
        let mockItems: any = require("./items.json");
        let searchResults = mockItems.filter((item) => {
            return item.Title.toLowerCase().search(
                query.toLowerCase()) !== -1; //|| item.SFDesc.toLowerCase().search(query.toLowerCase()) !== -1;
        });

        let groupedItems = Utils.groupBy(searchResults, item => item[groupBy]);
        let sortedItems = Utils.sortByKey(groupedItems);
        return new Promise<any>((resolve) => {
            setTimeout(() => {
                resolve(sortedItems);
            }, 2000);
        });
    }
}