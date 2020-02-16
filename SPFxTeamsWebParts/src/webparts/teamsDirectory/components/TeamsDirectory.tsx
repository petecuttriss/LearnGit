import * as React from 'react';
import styles from './TeamsDirectory.module.scss';
import { ITeamsDirectoryProps } from './ITeamsDirectoryProps';
import { escape, isEmpty } from '@microsoft/sp-lodash-subset';
import { ITeamsDirectoryState } from './ITeamsDirectoryState';
import { Nav, INavLinkGroup, INavLink } from 'office-ui-fabric-react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { IListItem } from '../../../models/IListItem';
import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { SearchBox } from 'office-ui-fabric-react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';

const FIELDSTOSELECT: string = "Title,SFLink,ContentType,SIAStatus,SFDesc,Function,Activity,Subactivity";
const GROUPBY: string = "Activity";
const FILTER: string = "SIAStatus eq '03 Active'";

export default class TeamsDirectory extends React.Component<ITeamsDirectoryProps, ITeamsDirectoryState> {

  /**
   * Constructor
   * @param props 
   */
  constructor(props: ITeamsDirectoryProps) {

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

  /**
   * Check the updated component props and if necessary get an updated result set.
   * @param prevProps 
   * @param prevState 
   */
  public componentDidUpdate(prevProps, prevState) {
    if (this.props !== prevProps) {
      if (this.props.siteUrl !== prevProps.siteUrl
        || this.props.listId !== prevProps.listId) {
        this.setState({
          loading: true,
          items: []
        });

        this._getItems();
      }
    }
  }

  /**
   * Render the component
   */
  public render(): React.ReactElement<ITeamsDirectoryProps> {
    const mandatoryFieldsConfigured = this._mandatoryFieldsConfigured();

    return (
      <div className={styles.teamsDirectory}>
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
                      />
                    }
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

  /**
   * Return the list items from the configured Workplace Catalogue list
   */
  private _getItems() {
    if (this.props.siteUrl && this.props.listId) {
      // Use the filter values to query the sites registry for items matching this filter
      this.props.listService.getItems(this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, FILTER).then((listItems) => {
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

  /**
   * When a search string is entered query the list items by title and return results
   * @param event 
   */
  private _onSearch(event) {
    if (event === undefined ||
      event === null ||
      event === "") {
      this.props.listService.getItems(this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, FILTER).then((listItems) => {
        this._mapItems(listItems, true, false);
      });
    }
    else {
      this.props.listService.searchItems(event, this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, FILTER).then((listItems) => {
        this._mapItems(listItems, false, true);
      });
    }
  }

  /**
   * When the search is cleared, get the default configured list items
   * @param event 
   */
  private _onSearchClear(event) {
    this.props.listService.getItems(this.props.siteUrl, this.props.listId, FIELDSTOSELECT, GROUPBY, FILTER).then((listItems) => {
      this._mapItems(listItems, true, false);
    });
  }

  /**
   * Takes the returned list items and converts them to INavLinkGroup and INavLink items
   * @param groupItems 
   * @param collapse 
   * @param isSearch 
   */
  private _mapItems(groupItems: any[], collapse: boolean, isSearch: boolean): void {
    let tmpGroup: INavLinkGroup[] = [];
    Object.keys(groupItems).map((key, i) => {
      let tmpLinks: INavLink[] = [];
      groupItems[key].map((groupedItem: IListItem, index: number) => {
        tmpLinks.push({ key: groupedItem.Title, name: groupedItem.Title, url: groupedItem.SFLink === null ? "" : groupedItem.SFLink.Url, target: this.props.openLinksInANewTab ? "_blank" : "_self", disabled: groupedItem.SIAStatus !== "03 Active" });
      });
      tmpGroup.push({ name: key === "null" ? "Unknown" : key, links: tmpLinks, collapseByDefault: collapse });
    });

    this.setState({
      error: null,
      items: tmpGroup,
      isSearch: isSearch
    });
  }

  /**
   * Returns whether all mandatory fields are configured or not.
   * If the fields aren't configured we render the PlaceHolder component.
   */
  private _mandatoryFieldsConfigured(): boolean {
    return !isEmpty(this.props.siteUrl) &&
      !isEmpty(this.props.listId);
  }

  /**
   * Calls the parent web part to open the property pane
   */
  private _configureWebPart(): void {
    // Context of the web part
    this.props.configureStartCallback();
  }
}
