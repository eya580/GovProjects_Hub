import { StatutProjet, Priorite } from "./Enums";

// Liste SharePoint : Projets
export interface IProject {
  Id?: number;
  Title: string;
  CodeProjet: string;
  Description?: string;
  ChefProjetId?: number;
  ChefProjet?: { Title: string; EMail: string; Id: number };
  SponsorId?: number;
  Sponsor?: { Title: string; EMail: string; Id: number };
  Direction?: string;
  StatutProjet: StatutProjet;
  DateDebutReference?: string;
  DateFinReference?: string;
  DateDebutPrevue?: string;
  DateFinPrevue?: string;
  DateDebutReelle?: string;
  DateFinReelle?: string;
  PourcentageAchevement?: number;
  Priorite: Priorite;
  Budget?: number;
  Actif: boolean;
  SiteProjetUrl?: { Url: string; Description?: string };
  CreePar?: string;
  ModifiePar?: string;
}
