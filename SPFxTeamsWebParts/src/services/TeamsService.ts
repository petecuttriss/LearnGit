import {
  ClientSideText,
  ClientSideWebpart,
  sp,
  ClientSidePage
} from "@pnp/sp";
import { PageContext } from "@microsoft/sp-page-context";
import { AadTokenProviderFactory } from "@microsoft/sp-http";
import { graph, Channel, Channels } from "@pnp/graph";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { ITeamsService } from "../models/ITeamsService";

export class TeamsService implements ITeamsService {

  public static readonly serviceKey: ServiceKey<ITeamsService> = ServiceKey.create<ITeamsService>('spfx-teams-web-parts:TeamsService', TeamsService);

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {

      const pageContext = serviceScope.consume(PageContext.serviceKey);
      const tokenProviderFactory = serviceScope.consume(AadTokenProviderFactory.serviceKey);

      // we need to "spoof" the context object with the parts we need for PnPjs
      sp.setup({
        spfxContext: {
          pageContext: pageContext,
        }
      });

      graph.setup({
        spfxContext: {
          aadTokenProviderFactory: tokenProviderFactory,
          pageContext: pageContext,
        }
      });
    });
  }

  public async getGroupId(): Promise<string> {

    var id: string = "";

    var props: any = await sp.web.select("AllProperties")
      .expand("AllProperties")
      .get();

    if (props.AllProperties["GroupId"] != null) {
      id = props.AllProperties["GroupId"];
    }
    return id;
  }

  public async getChannels(groupId: string): Promise<any[]> {

    var channels: any[] = [];

    channels = await graph.teams.getById(groupId).channels.get();

    return channels;
  }

  public async getTabsFromChannel(groupId: string, channelId: string): Promise<any[]> {

    var tabs: any[] = [];

    tabs = await graph.teams.getById(groupId).channels.getById(channelId)
      .tabs
      .get();

    return tabs;
  }

  public async getMyTeams(): Promise<any[]> {
    var myTeams: any[] = [];

    myTeams = await graph.me.joinedTeams.get();

    return myTeams;
  }

  public async getTeam(groupId: string): Promise<any> {
    let team: any;

    team = await graph.teams.getById(groupId).get();

    return team;
  }
}