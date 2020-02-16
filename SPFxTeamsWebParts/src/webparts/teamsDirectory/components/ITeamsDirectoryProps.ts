import { DisplayMode } from "@microsoft/sp-core-library";
import { IListService } from "../../../models/IListService";

export interface ITeamsDirectoryProps {
  title: string;
  siteUrl: string;
  listId: string;
  hideTitle: boolean;
  hideSearchBox: boolean;
  openLinksInANewTab: boolean;
  displayMode: DisplayMode;
  listService: IListService;
  updateProperty: (value: string) => void;
  configureStartCallback: () => void;
}
