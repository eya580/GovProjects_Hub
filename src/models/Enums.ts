// Enums métier communs - valeurs alignées sur les colonnes "Choix" des listes SharePoint

export enum StatutProjet {
  Brouillon = "Brouillon",
  Planifie = "Planifié",
  EnCours = "En cours",
  Suspendu = "Suspendu",
  Cloture = "Clôturé"
}

export enum Priorite {
  Faible = "Faible",
  Normale = "Normale",
  Haute = "Haute",
  Critique = "Critique"
}

export enum TypeActivite {
  Phase = "Phase",
  Lot = "Lot",
  Tache = "Tâche",
  Jalon = "Jalon"
}

export enum StatutActivite {
  NonDemarree = "Non démarrée",
  EnCours = "En cours",
  Terminee = "Terminée",
  Suspendue = "Suspendue"
}

export enum ModeCalculAvancement {
  Manuel = "Manuel",
  Charge = "Charge",
  Duree = "Durée"
}

export enum TypeLien {
  FinDebut = "Fin-Début",
  DebutDebut = "Début-Début",
  FinFin = "Fin-Fin",
  DebutFin = "Début-Fin"
}

export enum StatutAffectation {
  Prevue = "Prévue",
  Active = "Active",
  Terminee = "Terminée",
  Annulee = "Annulée"
}

export enum RoleProjet {
  Developpeur = "Développeur",
  Testeur = "Testeur",
  PM = "PM",
  BA = "BA",
  Autre = "Autre"
}

export enum MotifCapacite {
  Normal = "Normal",
  TempsPartiel = "Temps partiel",
  Formation = "Formation",
  Indisponibilite = "Indisponibilité"
}

export enum TypePeriode {
  Hebdomadaire = "Hebdomadaire",
  Mensuelle = "Mensuelle"
}

export enum StatutPeriode {
  Preparee = "Préparée",
  Ouverte = "Ouverte",
  Fermee = "Fermée"
}

export enum StatutFeuille {
  Brouillon = "Brouillon",
  Soumise = "Soumise",
  ACorriger = "À corriger",
  Validee = "Validée",
  Rejetee = "Rejetée"
}

export enum TypeLigne {
  Projet = "Projet",
  Administrative = "Administrative"
}

export enum StatutLigne {
  Brouillon = "Brouillon",
  Soumise = "Soumise",
  Validee = "Validée",
  Rejetee = "Rejetée"
}

export enum SourceSaisie {
  Web = "Web",
  Mobile = "Mobile",
  Import = "Import"
}

export enum TypeValeurParametre {
  Texte = "Texte",
  Nombre = "Nombre",
  Booleen = "Booléen",
  JSON = "JSON"
}

export enum TypeObjetHistorique {
  Projet = "Projet",
  Activite = "Activité",
  Affectation = "Affectation",
  FeuilleDeTemps = "Feuille de temps"
}

export enum ActionHistorique {
  Creation = "Création",
  Modification = "Modification",
  Soumission = "Soumission",
  Validation = "Validation",
  Rejet = "Rejet"
}
