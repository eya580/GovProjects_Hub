import { StatutAffectation } from "./Enums";

// Liste SharePoint : Affectations
export interface IAssignment {
  Id?: number;
  ProjetId: number;
  ActiviteId: number;
  RessourceId: number;
  DateDebutAffectation: string;
  DateFinAffectation: string;
  ChargeAffecteeHeures?: number;
  UniteAffectation?: number;
  HeuresParJour?: number;
  RoleProjet?: string;
  ResponsableTache: boolean;
  AutoriseSaisieTemps: boolean;
  StatutAffectation: StatutAffectation;
  Actif: boolean;
}

// Vue enrichie utilisée par le web part "Gestion des affectations"
export interface IAssignmentView extends IAssignment {
  RessourceNom?: string;
  ActiviteNom?: string;
  ProjetNom?: string;
  CapaciteRestanteHeures?: number;
  EstEnSurcharge?: boolean;
}
