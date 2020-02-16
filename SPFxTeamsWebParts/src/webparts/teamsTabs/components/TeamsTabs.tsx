import * as React from 'react';
import styles from './TeamsTabs.module.scss';
import { ITeamsTabsProps } from './ITeamsTabsProps';
import { Nav } from 'office-ui-fabric-react/lib/Nav';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';

export interface IReactTeamsTabsPnpjsState {
  error: string;
  pivotArray: any[];
}

export default class TeamsTabs extends React.Component<ITeamsTabsProps, IReactTeamsTabsPnpjsState> {

  /**
   * Constructor
   */
  constructor(props: ITeamsTabsProps) {
    super(props);

    this.state = {
      error: null,
      pivotArray: []
    };
  }

  /**
   * componentDidMount
   * 
   * Use the TeamsService to get the current site's groupId.
   * Use the groupId to query Teams for the associated Channels.
   * For each channel query Teams for the associated Tabs.
   * Requires API Permissions.  Need to add a check and respond with a nice
   * message if the permissions haven't been approved.
   */
  public componentDidMount() {
    var groupId: Promise<string> = this.props.teamsService.getGroupId();

    groupId.then(group => {
      console.log("GroupID: " + group);
      var tmpChannels: any[] = [];
      if (group != "") {
        var channels: Promise<any[]> = this.props.teamsService.getChannels(group);

        channels.then(chans => {
          console.log("Channels " + chans.length);
          chans.forEach(channel => {
            var tabs: Promise<any[]> = this.props.teamsService.getTabsFromChannel(group, channel.id);
            var tmpTabs: any[] = [];
            tabs.then(itemTabs => {
              console.log("Channel" + channel.displayName + "tabs " + itemTabs.length);
              itemTabs.forEach(tab => {
                tmpTabs.push({ key: tab.id, name: decodeURI(tab.displayName), url: tab.webUrl, target: '_blank' });
              });
              tmpChannels.push({ name: channel.displayName + " (" + tmpTabs.length + ")", links: tmpTabs, collapseByDefault: true });
              tmpChannels.sort(this.mySorter);
              this.setState({ pivotArray: tmpChannels });
            });
          });
        });

      } else {
        this.setState({ error: "No Teams instance associated with this site." });
      }
    });


  }
  public render(): React.ReactElement<ITeamsTabsProps> {
    return (
      <div className={styles.teamsTabs}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={'ms-Grid'} dir="ltr">
                <div className={'ms-Grid-row'}>
                  <div className='ms-Grid-col ms-sm12'>
                    <WebPartTitle className={styles.title} displayMode={this.props.displayMode}
                      title={this.props.title}
                      updateProperty={this.props.updateProperty} />
                    {this.state.error &&
                      <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>
                    }
                    <Nav
                      groups={this.state.pivotArray}
                      className={styles.nav}
                      key={Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15)}
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private mySorter(a: any, b: any) {
    var x = a.name.toLowerCase();
    var y = b.name.toLowerCase();
    //fix to manage general channel at first position, like Teams order
    //verify language general label
    if (x.startsWith("general")) {
      return -1;
    } else if (y.startsWith("general")) {
      return 1;
    }
    return ((x < y) ? -1 : ((x > y) ? 1 : 0));
  }
}
