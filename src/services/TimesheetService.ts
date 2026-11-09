import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP, LISTS, mapLookupIds, buildLookupWritePayload, filtreLookupEgal } from "./SPConfig";
import { ITimesheet, ITimesheetLine, ITimesheetPeriod } from "../models/ITimesheet";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { StatutFeuille } from "../models/Enums";

// Colonnes Lookup (titrées avec le suffixe "Id") sur FeuillesTemps et LignesFeuilleTemps.
// NB : "Utilisateur", "SoumisPar", "Validateur" sont des champs Personne classiques
// (nommés SANS suffixe "Id"), donc leur champ ombre XxxId fonctionne nativement - pas de correctif requis.
const CHAMPS_LOOKUP_FEUILLE = ["RessourceId", "PeriodeId"];
const CHAMPS_LOOKUP_LIGNE = ["FeuilleTempsId", "RessourceId", "ProjetId", "ActiviteId", "TypeTacheAdministrativeId", "AffectationId"];

const SELECT_FEUILLE = [
  "Id", "Title", "CodeFeuilleTemps", "RessourceId/Id", "UtilisateurId", "PeriodeId/Id",
  "StatutFeuille", "TotalHeures", "TotalHeuresProjet", "TotalHeuresAdministratives",
  "DateSoumission", "SoumisParId", "ValidateurId", "DateValidation",
  "CommentaireRessource", "CommentaireValidateur", "Verrouillee"
];

const SELECT_LIGNE = [
  "Id", "FeuilleTempsId/Id", "RessourceId/Id", "DateTravail", "TypeLigne",
  "ProjetId/Id", "ActiviteId/Id", "TypeTacheAdministrativeId/Id", "AffectationId/Id",
  "NombreHeures", "Commentaire", "StatutLigne", "Facturable",
  "DateCreationSaisie", "SourceSaisie", "Modifiable"
];

export class TimesheetService {
  constructor(private context: WebPartContext) {}

  public async getAllPeriodes(): Promise<ITimesheetPeriod[]> {
    const sp = getSP(this.context);
    const results: ITimesheetPeriod[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.PeriodesFeuilleTemps).items
      .orderBy("DateDebut", false).top(500)) {
      results.push(...(batch as ITimesheetPeriod[]));
    }

    return results;
  }

  // --- Périodes ---

  public async getPeriodesOuvertes(): Promise<ITimesheetPeriod[]> {
    const sp = getSP(this.context);
    const results: ITimesheetPeriod[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.PeriodesFeuilleTemps).items
      .filter("StatutPeriode eq 'Ouverte' and Actif eq 1")
      .orderBy("DateDebut", false).top(500)) {
      results.push(...(batch as ITimesheetPeriod[]));
    }

    return results;
  }

  // --- Feuilles de temps (en-tête) ---

  public async getOrCreateFeuille(ressourceId: number, periodeId: number, utilisateurId: number): Promise<ITimesheet> {
    const sp = getSP(this.context);
    const existantes = await sp.web.lists.getByTitle(LISTS.FeuillesTemps).items
      .select(...SELECT_FEUILLE)
      .expand("RessourceId", "PeriodeId")
      .filter(`${filtreLookupEgal("RessourceId", ressourceId)} and ${filtreLookupEgal("PeriodeId", periodeId)}`)
      .top(1)();

    if (existantes.length > 0) { return mapLookupIds(existantes[0], CHAMPS_LOOKUP_FEUILLE) as ITimesheet; }

    const nouvelle: Record<string, unknown> = buildLookupWritePayload({
      Title: `FT-${ressourceId}-${periodeId}`,
      CodeFeuilleTemps: `FT-${ressourceId}-${periodeId}-${Date.now()}`,
      RessourceId: ressourceId,
      UtilisateurId: utilisateurId,
      PeriodeId: periodeId,
      StatutFeuille: StatutFeuille.Brouillon,
      TotalHeures: 0,
      TotalHeuresProjet: 0,
      TotalHeuresAdministratives: 0,
      Verrouillee: false
    }, CHAMPS_LOOKUP_FEUILLE);

    const result = await sp.web.lists.getByTitle(LISTS.FeuillesTemps).items.add(nouvelle);
    return result as unknown as ITimesheet;
  }

  public async getFeuillesEnAttente(): Promise<ITimesheet[]> {
    const sp = getSP(this.context);
    const results: ITimesheet[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.FeuillesTemps).items
      .select(...SELECT_FEUILLE)
      .expand("RessourceId", "PeriodeId")
      .filter(`StatutFeuille eq '${StatutFeuille.Soumise}'`)
      .orderBy("DateSoumission", true).top(500)) {
      results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_FEUILLE) as unknown as ITimesheet));
    }

    return results;
  }

  // --- Lignes ---

  public async getLignes(feuilleTempsId: number): Promise<ITimesheetLine[]> {
    const sp = getSP(this.context);
    const results: ITimesheetLine[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.LignesFeuilleTemps).items
      .select(...SELECT_LIGNE)
      .expand("FeuilleTempsId", "RessourceId", "ProjetId", "ActiviteId", "TypeTacheAdministrativeId", "AffectationId")
      .filter(filtreLookupEgal("FeuilleTempsId", feuilleTempsId))
      .top(500)) {
      results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_LIGNE) as unknown as ITimesheetLine));
    }

    return results;
  }

  public async upsertLigne(ligne: Partial<ITimesheetLine>): Promise<ITimesheetLine> {
    const sp = getSP(this.context);
    const payload = buildLookupWritePayload(ligne as Record<string, unknown>, CHAMPS_LOOKUP_LIGNE);
    if (ligne.Id) {
      await sp.web.lists.getByTitle(LISTS.LignesFeuilleTemps).items.getById(ligne.Id).update(payload);
      return ligne as ITimesheetLine;
    }
    const result = await sp.web.lists.getByTitle(LISTS.LignesFeuilleTemps).items.add(payload);
    return result as unknown as ITimesheetLine;
  }

  // Retourne le total des heures administratives saisies par une ressource sur une plage de dates
  // (utilisé par CapaciteCharge pour déduire les indisponibilités réelles de la capacité nette)
  public async getHeuresAdminParRessource(ressourceId: number, dateDebut: string, dateFin: string): Promise<number> {
    const sp = getSP(this.context);
    let total = 0;

    for await (const batch of sp.web.lists.getByTitle(LISTS.LignesFeuilleTemps).items
      .select("NombreHeures")
      .filter(`${filtreLookupEgal("RessourceId", ressourceId)} and TypeLigne eq 'Administrative' and DateTravail ge datetime'${dateDebut}T00:00:00Z' and DateTravail le datetime'${dateFin}T23:59:59Z'`)
      .top(500)) {
      total += (batch as { NombreHeures?: number }[]).reduce((s, l) => s + (l.NombreHeures || 0), 0);
    }

    return total;
  }

  public async deleteLigne(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.LignesFeuilleTemps).items.getById(id).delete();
  }

  // --- Règles de gestion / Workflow ---

  // Règle d'intégrité : une ligne pointe vers une affectation projet OU une tâche administrative, jamais les deux
  public validerIntegriteLigne(ligne: Partial<ITimesheetLine>): boolean {
    const aAffectation = !!ligne.AffectationId;
    const aTacheAdmin = !!ligne.TypeTacheAdministrativeId;
    return aAffectation !== aTacheAdmin;
  }

  public calculerTotaux(lignes: ITimesheetLine[]): { total: number; projet: number; administratif: number } {
    const projet = lignes.filter(l => l.TypeLigne === "Projet").reduce((s, l) => s + (l.NombreHeures || 0), 0);
    const administratif = lignes.filter(l => l.TypeLigne === "Administrative").reduce((s, l) => s + (l.NombreHeures || 0), 0);
    return { total: projet + administratif, projet, administratif };
  }

  // Flux 3 - Soumission : contrôle des heures/jours/tâches autorisées puis passage à "Soumise"
  public async soumettre(feuille: ITimesheet, lignes: ITimesheetLine[]): Promise<void> {
    if (feuille.Verrouillee) { throw new Error("La feuille est verrouillée et ne peut plus être modifiée."); }
    if (lignes.some(l => !this.validerIntegriteLigne(l))) {
      throw new Error("Une ligne doit pointer vers une tâche projet OU une tâche administrative, pas les deux.");
    }
    const totaux = this.calculerTotaux(lignes);
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.FeuillesTemps).items.getById(feuille.Id as number).update({
      StatutFeuille: StatutFeuille.Soumise,
      DateSoumission: new Date().toISOString(),
      TotalHeures: totaux.total,
      TotalHeuresProjet: totaux.projet,
      TotalHeuresAdministratives: totaux.administratif
    });
  }

  // Flux 4 - Validation : verrouille la feuille et alimente la charge réelle des tâches
  public async valider(feuilleId: number, validateurId: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.FeuillesTemps).items.getById(feuilleId).update({
      StatutFeuille: StatutFeuille.Validee,
      ValidateurId: validateurId,
      DateValidation: new Date().toISOString(),
      Verrouillee: true
    });
  }

  // Flux 5 - Rejet : passage à "À corriger", commentaire, réactivation de la saisie
  public async rejeter(feuilleId: number, commentaireValidateur: string): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.FeuillesTemps).items.getById(feuilleId).update({
      StatutFeuille: StatutFeuille.ACorriger,
      CommentaireValidateur: commentaireValidateur,
      Verrouillee: false
    });
  }
}