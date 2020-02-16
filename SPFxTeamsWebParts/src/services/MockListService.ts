import { IListService } from "../models/IListService";
import { Utils } from "./Utils";

export class MockListService implements IListService {
    public async getPropertyBagItem(item: string): Promise<string> {
        var propertyBagItem: string = "Water Supply";

        return propertyBagItem;
    }

    public getItems(siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any> {
        let mockItems: any = require("./items.json");
        let groupedItems = Utils.groupBy(mockItems, item => item[groupBy]);
        let sortedItems = Utils.sortByKey(groupedItems);

        return new Promise<any>((resolve) => {
            //setTimeout(() => {
            resolve(sortedItems);
            //}, 2000);
        });
    }

    public searchItems(query: string, siteUrl: string, listId: string, groupBy: string, filter?: string): Promise<any> {
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