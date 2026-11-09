// Liste SharePoint : HistoriqueActions
export interface IAuditHistory {
  Id?: number;
  TypeObjet: string;
  ObjetId: number;
  Action: string;
  AncienneValeur?: string;
  NouvelleValeur?: string;
  UtilisateurId?: number;
  DateAction: string;
  Commentaire?: string;
}

// Liste SharePoint : ParametresApplication
export interface IApplicationSetting {
  Id?: number;
  Cle: string;
  Valeur: string;
  TypeValeur: string;
  Description?: string;
  Actif: boolean;
}
