import * as React from 'react';
import styles from './TeamsByActivityContainer.module.scss';
import { ITeamsByActivityContainerProps } from './ITeamsByActivityContainerProps';
import { ITeamsByActivityContainerState } from './ITeamsByActivityContainerState';
import { escape, isEmpty } from '@microsoft/sp-lodash-subset';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { DisplayMode } from '@microsoft/sp-core-library';
import { stringIsNullOrEmpty } from '@pnp/common';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { IListItem } from "../../../../models/IListItem";
import { ListService } from "../../../../services/ListService";
import { Nav, INavLinkGroup, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

const FIELDSTOSELECT: string = "Title,SFLink,ContentType,SIAStatus,SFDesc,Function,Activity,Subactivity";
const GROUPBY: string = "Subactivity";

export default class ActivityTeamsContainer extends React.Component<ITeamsByActivityContainerProps, ITeamsByActivityContainerState> {

  /**
   * Constructor
   */
  constructor(props: ITeamsByActivityContainerProps) {
    super(props);

    this._onSearch = this._onSearch.bind(this);
    this._onSearchClear = this._onSearchClear.bind(this);
    this._configureWebPart = this._configureWebPart.bind(this);

    this.state = {
      error: null,
      loading: true,
      isSearch: false,
      items: []
    };
  }

  /**
   * componentDidMount
   * 
   * Get the list items from the sites register.
   */
  public componentDidMount() {
    this._getItems();
  }

  public componentDidUpdate(prevProps, prevState) {
    if (this.props !== prevProps) {
      if (this.props.siteUrl !== prevProps.siteUrl
        || this.props.listId !== prevProps.listId
        || this.props.odataFilter !== prevProps.odataFilter) {
        this.setState({
          loading: true,
          items: []
        });

        this._getItems();
      }
    }
  }

  /**
   * render
   * 
   * Create an navigation component.
   * On first load, display the 'loading...' spinner.
   * If any errors, display the 'error details' panel.
   * Once the data has come back and there are no errors, display the collapsed nav.
   */
  public render(): React.ReactElement<ITeamsByActivityContainerProps> {
    const mandatoryFieldsConfigured = this._mandatoryFieldsConfigured();

    return (
      <div className={styles.activityTeams}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <div className={'ms-Grid'} dir="ltr">
                <div className={'ms-Grid-row'}>
                  {!mandatoryFieldsConfigured &&
                    <Placeholder iconName='Edit'
                      iconText='Configure your web part'
                      description='Please configure the web part.'
                      buttonLabel='Configure'
                      hideButton={this.props.displayMode === DisplayMode.Read}
                      onConfigure={this._configureWebPart} />
                  }
                  {mandatoryFieldsConfigured && !this.props.hideTitle &&
                    <div className='ms-Grid-col ms-sm12 ms-md7'>
                      <WebPartTitle className={styles.title} displayMode={this.props.displayMode}
                        title={this.props.title}
                        updateProperty={this.props.updateProperty} />
                    </div>
                  }
                  {mandatoryFieldsConfigured && !this.props.hideSearchBox &&
                    <div className='ms-Grid-col ms-sm12 ms-md5'>
                      <SearchBox onSearch={this._onSearch} onClear={this._onSearchClear} className={styles.searchbox} placeholder="Search by title" />
                    </div>
                  }
                </div>
                <div className={'ms-Grid-row'}>
                  <div className='ms-Grid-col ms-sm12'>
                    {mandatoryFieldsConfigured && this.state.error &&
                      <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>
                    }
                    {mandatoryFieldsConfigured && Object.keys(this.state.items).length > 0 && !this.state.error &&
                      <Nav
                        groups={this.state.items}
                        collapsedStateText="Collapsed"
                        expandedStateText="Expand"
                        className={styles.nav} 
                        key={Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15)}
                      />}
                    {mandatoryFieldsConfigured && Object.keys(this.state.items).length == 0 && !this.state.error &&
                      <MessageBar messageBarType={MessageBarType.info}>No results found.</MessageBar>
                    }
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _getItems() {
    if (this.props.siteUrl && this.props.listId && this.props.listService) {
      // Use the filter values to query the sites registry for items matching this filter
      this.props.listService.getItems(this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, this.props.odataFilter).then((listItems) => {
        if (listItems.error) {
          this.setState({
            error: listItems.error,
            items: [],
            isSearch: false
          });
        } else {
          this._mapItems(listItems, true, false);
        }
      });
    }
  }

  private _onSearch(event) {
    if (event === undefined ||
      event === null ||
      event === "") {
        this.props.listService.getItems(this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, this.props.odataFilter).then((listItems) => {
        this._mapItems(listItems, true, false);
      });
    }
    else {
      this.props.listService.searchItems(event, this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, this.props.odataFilter).then((listItems) => {
        this._mapItems(listItems, false, true);
      });
    }
  }

  private _onSearchClear(event) {
    this.props.listService.getItems(this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, this.props.odataFilter).then((listItems) => {
      this._mapItems(listItems, true, false);
    });
  }

  private _mapItems(groupItems: any[], collapse: boolean, isSearch: boolean) {
    let tmpGroup: INavLinkGroup[] = [];
    Object.keys(groupItems).map((key, i) => {
      let tmpLinks: INavLink[] = [];
      groupItems[key].map((groupedItem: IListItem, index: number) => {
        tmpLinks.push({ key: Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15), name: groupedItem.Title, url: groupedItem.SFLink === null ? "" : groupedItem.SFLink.Url, target: this.props.openLinksInANewTab ? "_blank" : "_self", disabled: groupedItem.SIAStatus !== "03 Active" });
      });
      tmpGroup.push({ name: key === "null" ? "Unknown" : key, links: tmpLinks, collapseByDefault: collapse, automationId: Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15)});
    });

    this.setState({
      error: null,
      items: tmpGroup,
      isSearch: isSearch
    });
  }

  /*************************************************************************************
   * Returns whether all mandatory fields are configured or not
   *************************************************************************************/
  private _mandatoryFieldsConfigured(): boolean {
    return !isEmpty(this.props.siteUrl) &&
      !isEmpty(this.props.listId);
  }

  private _configureWebPart() {
    // Context of the web part
    this.props.configureStartCallback();
  }
}
