export interface ILinkItem {
    Description: string;
    Url: string;
}
export interface IListItem {
    Id: number;
    Title: string;
    SFDesc: string;
    SIAStatus: string;
    SFLink: ILinkItem;
    Function: string;
    Activity: string;
    Subactivity: string;
}
export interface IGroupedItems {
    Title: string;
    Items: IListItem[];
}