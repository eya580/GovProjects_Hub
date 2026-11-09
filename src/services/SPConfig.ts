import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// Noms EXACTS des listes SharePoint - à conserver tels quels (cf. document de cadrage)
export const LISTS = {
  Projets: "Projets",
  Activites: "Activites",
  PredecesseursActivites: "PredecesseursActivites",
  Ressources: "Ressources",
  CapacitesRessources: "CapacitesRessources",
  Affectations: "Affectations",
  TypesTachesAdministratives: "TypesTachesAdministratives",
  PeriodesFeuilleTemps: "PeriodesFeuilleTemps",
  FeuillesTemps: "FeuillesTemps",
  LignesFeuilleTemps: "LignesFeuilleTemps",
  HistoriqueActions: "HistoriqueActions",
  ParametresApplication: "ParametresApplication"
} as const;

let _sp: SPFI;

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === undefined && context !== undefined) {
    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};

// --- Helpers génériques pour les colonnes Lookup ---
// De nombreuses colonnes de liaison (RessourceId, ProjetId, ActiviteId, PeriodeId,
// FeuilleTempsId, AffectationId, TypeTacheAdministrativeId, PredecesseurId, ...) sont des
// colonnes SharePoint de type Lookup nommées littéralement avec le suffixe "Id".
// Pour SharePoint, un tel champ est une "navigation property" : on ne peut donc pas
// le lire/filtrer comme un simple nombre, ni l'écrire tel quel.
//
// Lecture : select("Champ/Id") + expand("Champ")   -> renvoie { Champ: { Id: number } }
// Filtre  : "Champ/Id eq 5"                          (jamais "Champ eq 5")
// Écriture: { ChampId: 5 }                           (jamais { Champ: 5 })

// Construit les fragments select/expand nécessaires pour un ensemble de champs lookup
export const lookupSelectExpand = (champsLookup: string[]): { select: string[]; expand: string[] } => ({
  select: champsLookup.map(c => `${c}/Id`),
  expand: [...champsLookup]
});

// Transforme un item brut PnPjs ({ Champ: { Id: number } }) en item "métier" ({ Champ: number })
export function mapLookupIds<T extends Record<string, unknown>>(item: T, champsLookup: string[]): T {
  const result: Record<string, unknown> = { ...item };
  champsLookup.forEach(c => {
    const valeur = result[c] as { Id?: number } | undefined;
    result[c] = valeur?.Id;
  });
  return result as T;
}

// Transforme un payload métier ({ Champ: number }) en payload d'écriture SharePoint ({ ChampId: number })
export function buildLookupWritePayload<T extends Record<string, unknown>>(payload: T, champsLookup: string[]): Record<string, unknown> {
  const result: Record<string, unknown> = { ...payload };
  champsLookup.forEach(c => {
    if (result[c] !== undefined) {
      result[`${c}Id`] = result[c];
      delete result[c];
    }
  });
  return result;
}

// Construit une clause $filter sûre pour un champ lookup : "ChampId eq 5"
// Les noms passés incluent déjà le suffixe "Id" (ex: "RessourceId") — on ne l'ajoute pas.
export const filtreLookupEgal = (champ: string, id: number): string => `${champ} eq ${id}`;