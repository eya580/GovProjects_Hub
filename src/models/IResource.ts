// Liste SharePoint : Ressources
export interface IResource {
  Id?: number;
  Title: string;
  UtilisateurId?: number;
  Utilisateur?: { Title: string; EMail: string; Id: number; LoginName?: string };
  Email?: string;
  Matricule?: string;
  Profil?: string;
  Departement?: string;
  ManagerId?: number;
  Manager?: { Title: string; EMail: string; Id: number };
  CapaciteJournaliereHeures?: number;
  CapaciteHebdomadaireHeures?: number;
  TauxAffectationMaximum?: number;
  DateEntree?: string;
  DateSortie?: string;
  Actif: boolean;
  PeutSaisirTemps: boolean;
  PeutGererPlanning: boolean;
}

// Liste SharePoint : CapacitesRessources
export interface ICapacity {
  Id?: number;
  RessourceId: number;
  DateDebut: string;
  DateFin: string;
  CapaciteHeuresParJour: number;
  TauxDisponibilite: number;
  CapaciteTotaleHeures: number;
  Motif: string;
  Actif: boolean;
}
