export interface ITeamsService {
    getGroupId(): Promise<string>;
    getChannels(groupId: string): Promise<any[]>;
    getTabsFromChannel(groupId: string, channelId: string): Promise<any[]>;
    getMyTeams(): Promise<any[]>;
    getTeam(groupId: string): Promise<any>;
}