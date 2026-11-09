import { StatutFeuille, TypeLigne, StatutLigne, SourceSaisie } from "./Enums";

// Liste SharePoint : PeriodesFeuilleTemps
export interface ITimesheetPeriod {
  Id?: number;
  Title: string;
  CodePeriode: string;
  TypePeriode: string;
  DateDebut: string;
  DateFin: string;
  DateLimiteSoumission?: string;
  StatutPeriode: string;
  Annee: number;
  NumeroSemaine?: number;
  Mois?: number;
  Actif: boolean;
}

// Liste SharePoint : FeuillesTemps (en-tête)
export interface ITimesheet {
  Id?: number;
  Title: string;
  CodeFeuilleTemps: string;
  RessourceId: number;
  RessourceIdId?: number;
  UtilisateurId?: number;
  UtilisateurIdId?: number;
  PeriodeId: number;
  PeriodeIdId?: number;
  StatutFeuille: StatutFeuille;
  TotalHeures?: number;
  TotalHeuresProjet?: number;
  TotalHeuresAdministratives?: number;
  DateSoumission?: string;
  SoumisParId?: number;
  ValidateurId?: number;
  DateValidation?: string;
  CommentaireRessource?: string;
  CommentaireValidateur?: string;
  Verrouillee: boolean;
}

// Liste SharePoint : LignesFeuilleTemps
export interface ITimesheetLine {
  Id?: number;
  FeuilleTempsId: number;
  FeuilleTempsIdId?: number;
  RessourceId: number;
  RessourceIdId?: number;
  DateTravail: string;
  TypeLigne: TypeLigne;
  ProjetId?: number;
  ProjetIdId?: number;
  ActiviteId?: number;
  ActiviteIdId?: number;
  TypeTacheAdministrativeId?: number;
  TypeTacheAdministrativeIdId?: number;
  AffectationId?: number;
  AffectationIdId?: number;
  NombreHeures: number;
  Commentaire?: string;
  StatutLigne: StatutLigne;
  Facturable?: boolean;
  DateCreationSaisie?: string;
  SourceSaisie: SourceSaisie;
  Modifiable: boolean;
}

// Liste SharePoint : TypesTachesAdministratives
export interface IAdminTaskType {
  Id?: number;
  Title: string;
  CodeType: string;
  Description?: string;
  Couleur?: string;
  NecessiteValidation: boolean;
  EstAbsence: boolean;
  ImpacteCapacite: boolean;
  SaisieHeuresAutorisee: boolean;
  SaisieJourneeAutorisee: boolean;
  Actif: boolean;
  Ordre?: number;
}

// Vue journalière utilisée par le composant "Ma feuille de temps"
export interface ITimesheetDayCell {
  date: string;
  ligne?: ITimesheetLine;
  heures: number;
}

export interface ITimesheetRow {
  cle: string; // AffectationId ou TypeTacheAdministrativeId
  libelle: string;
  type: TypeLigne;
  affectationId?: number;
  typeTacheAdministrativeId?: number;
  jours: { [date: string]: ITimesheetDayCell };
  totalLigne: number;
}