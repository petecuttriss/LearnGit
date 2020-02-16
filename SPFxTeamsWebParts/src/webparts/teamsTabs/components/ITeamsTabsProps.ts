import { DisplayMode } from "@microsoft/sp-core-library";
import { ITeamsService } from "../../../models/ITeamsService";

export interface ITeamsTabsProps {
  title: string;
  displayMode: DisplayMode;
  teamsService: ITeamsService;
  updateProperty: (value: string) => void;
}
