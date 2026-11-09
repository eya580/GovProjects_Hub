import { TypeActivite, StatutActivite, ModeCalculAvancement } from "./Enums";

// Liste SharePoint : Activites (Phases / Lots / Tâches / Jalons)
// Forme "métier" utilisée partout côté UI/services : les lookups (ProjetId, ParentActiviteId)
// sont déjà résolus en simples nombres (voir ActivityService.mapLookups).
export interface IActivity {
  Id?: number;
  Title: string;
  CodeActivite: string;
  ProjetId: number;
  TypeActivite: TypeActivite;
  ParentActiviteId?: number;
  WBS?: string;
  NiveauWBS?: number;
  Ordre?: number;
  Description?: string;
  DateDebutReference?: string;
  DateFinReference?: string;
  DateDebutPrevue?: string;
  DateFinPrevue?: string;
  DateDebutReelle?: string;
  DateFinReelle?: string;
  PourcentageAchevement?: number;
  ChargePlanifieeHeures?: number;
  ChargeReelleHeures?: number;
  DureePlanifieeJours?: number;
  StatutActivite: StatutActivite;
  EstJalon: boolean;
  AutoriseSaisieTemps: boolean;
  EstActive: boolean;
  ResponsableActiviteId?: number;
  ModeCalculAvancement: ModeCalculAvancement;
  ModifiableParRessource: boolean;
}

// Noeud enrichi pour l'affichage hiérarchique WBS (calculé côté client)
export interface IActivityNode extends IActivity {
  children: IActivityNode[];
  depth: number;
}

// Liste SharePoint : PredecesseursActivites
export interface IActivityPredecessor {
  Id?: number;
  ActiviteId: number;
  PredecesseurId: number;
  TypeLien: string;
  DecalageJours?: number;
  ProjetId: number;
  Actif: boolean;
}

// --- Formes "brutes" telles que renvoyées par PnPjs (lookups expand/select) ---
// Utilisées uniquement à l'intérieur d'ActivityService, jamais exposées à l'UI.

export interface IActivitySP extends Omit<IActivity, "ProjetId" | "ParentActiviteId"> {
  ProjetId: { Id: number };
  ParentActiviteId?: { Id: number };
}

export interface IActivityPredecessorSP extends Omit<IActivityPredecessor, "ProjetId" | "ActiviteId" | "PredecesseurId"> {
  ProjetId:       { Id: number };
  ActiviteId:     { Id: number };
  PredecesseurId: { Id: number };
}

// Payload d'écriture PnPjs : les champs lookup métier (ProjetId, ParentActiviteId)
// sont remplacés par leurs équivalents FieldNameId attendus par l'API SharePoint.
export interface IActivityWritePayload extends Partial<Omit<IActivity, "ProjetId" | "ParentActiviteId">> {
  ProjetIdId?: number;
  ParentActiviteIdId?: number;
}