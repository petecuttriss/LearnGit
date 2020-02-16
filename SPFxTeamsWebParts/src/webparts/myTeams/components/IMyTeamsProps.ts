import { DisplayMode } from "@microsoft/sp-core-library";
import { ITeamsService } from "../../../models/ITeamsService";

export interface IMyTeamsProps {
  title: string;
  displayMode: DisplayMode;
  teamsService: ITeamsService;
  updateProperty: (value: string) => void;
}
