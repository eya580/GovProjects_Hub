import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IProject } from "../../../../models/IProject";

export interface IPortefeuilleProjetsProps {
  context: WebPartContext;
  titre: string;
  onOpenProjet?: (projet: IProject) => void;
}
