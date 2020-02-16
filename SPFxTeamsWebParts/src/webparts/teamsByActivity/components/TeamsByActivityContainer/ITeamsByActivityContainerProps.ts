import { DisplayMode } from '@microsoft/sp-core-library';
import { IListService } from '../../../../models/IListService';

export interface ITeamsByActivityContainerProps {
  title: string;
  siteUrl: string;
  listId: string;
  odataFilter: string;
  hideTitle: boolean;
  hideSearchBox: boolean;
  openLinksInANewTab: boolean;
  displayMode: DisplayMode;
  listService: IListService;
  updateProperty: (value: string) => void;
  configureStartCallback: () => void;
}