export interface IListService {
    getPropertyBagItem(item:string): Promise<string>;
    getItems(siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any>;
    searchItems(query: string, siteUrl: string, listId: string, fields: string, groupBy: string, filter?: string): Promise<any>;
}