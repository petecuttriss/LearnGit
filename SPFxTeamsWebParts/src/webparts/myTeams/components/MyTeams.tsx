import * as React from 'react';
import styles from './MyTeams.module.scss';
import { IMyTeamsProps } from './IMyTeamsProps';
import { Nav, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';


export interface IMyTeamsState {
  error: string;
  teams: any[];
}

export default class MyTeams extends React.Component<IMyTeamsProps, IMyTeamsState> {
  /**
 * Constructor
 */
  constructor(props: IMyTeamsProps) {
    super(props);

    this.state = {
      error: null,
      teams: []
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
    this._getMyTeams();
  }

  public componentDidUpdate(prevProps, prevState) {
    if (this.props !== prevProps) {
        this._getMyTeams();
    }
  }

  /**
   * _getMyTeams
   * 
   * Uses the Teams service to get a list
   */
  private _getMyTeams(): void {
    var myTeams: Promise<any[]> = this.props.teamsService.getMyTeams();

    myTeams.then(teams => {
      if(teams.length != 0){
        var teamTabs: any[] = [];
        var teamTabLinks: any[] = [];
        let promises = [];
        teams.forEach(team => {
          // Alternative would be to make this call during the onLinkClick call back.
          promises.push(this.props.teamsService.getTeam(team.id).then(teamProperties => {
            let webUrl = teamProperties.webUrl;
            teamTabLinks.push({ key: team.id, name: decodeURI(team.displayName), url: webUrl, target: '_blank', isExpanded: true });
          }));
        });

        Promise.all(promises).then(() => {
          teamTabLinks.sort(this.mySorter);
          teamTabs.push({ name: "My Teams (" + teams.length + ")", links: teamTabLinks, collapseByDefault: false });
          this.setState({ teams: teamTabs });
        });
      }else{
        this.setState({error: "You are not a direct member of any teams.", teams:[]});
      }
    });
  }

  public render(): React.ReactElement<IMyTeamsProps> {
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
                      groups={this.state.teams}
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
    return ((x < y) ? -1 : ((x > y) ? 1 : 0));
  }
}
