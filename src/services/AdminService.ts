import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { getSP, LISTS } from "./SPConfig";
import { IAdminTaskType, ITimesheetPeriod } from "../models/ITimesheet";
import { IApplicationSetting, IAuditHistory } from "../models/IAdministration";
import { StatutPeriode, ActionHistorique } from "../models/Enums";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export class AdminService {
  constructor(private context: WebPartContext) {}

  // --- Utilisateur connecté (pour l'historisation) ---

  public async getUtilisateurConnecte(): Promise<{ id: number; title: string; email: string }> {
    const sp = getSP(this.context);
    const currentUser = await sp.web.currentUser();
    return { id: currentUser.Id, title: currentUser.Title, email: currentUser.Email };
  }

  // --- Types de tâches administratives ---

  public async getTypesTachesAdministratives(actifSeulement = true): Promise<IAdminTaskType[]> {
    const sp = getSP(this.context);
    let query = sp.web.lists.getByTitle(LISTS.TypesTachesAdministratives).items.orderBy("Ordre", true);
    if (actifSeulement) { query = query.filter("Actif eq 1"); }

    const results: IAdminTaskType[] = [];
    for await (const batch of query.top(500)) {
      results.push(...(batch as IAdminTaskType[]));
    }

    return results;
  }

  private sanitizeAdminTaskType(type: Partial<IAdminTaskType>): Record<string, unknown> {
    const allowedFields: Array<keyof IAdminTaskType> = [
      "Title", "CodeType", "Description", "Couleur", "NecessiteValidation",
      "EstAbsence", "ImpacteCapacite", "SaisieHeuresAutorisee", "SaisieJourneeAutorisee",
      "Actif", "Ordre"
    ];
    const payload: Record<string, unknown> = {};
    for (const key of allowedFields) {
      const value = type[key];
      if (value === undefined || value === null) { continue; }
      if (typeof value === "number" && Number.isNaN(value)) { continue; }
      payload[key] = value;
    }
    return payload;
  }

  private extractAddedItemId(result: unknown): number {
    const typed = result as { data?: { Id?: number }; item?: { Id?: number } };
    return typed.data?.Id ?? typed.item?.Id ?? 0;
  }

  public async saveTypeTacheAdministrative(type: Partial<IAdminTaskType>): Promise<void> {
    const sp = getSP(this.context);
    const payload = this.sanitizeAdminTaskType(type);
    if (type.Id) {
      await sp.web.lists.getByTitle(LISTS.TypesTachesAdministratives).items.getById(type.Id).update(payload);
      await this.safeTracerAction({
        TypeObjet: "Type de tâche administrative",
        ObjetId: type.Id,
        Action: ActionHistorique.Modification,
        NouvelleValeur: JSON.stringify(payload)
      });
    } else {
      const result = await sp.web.lists.getByTitle(LISTS.TypesTachesAdministratives).items.add(payload);
      const createdId = this.extractAddedItemId(result);
      await this.safeTracerAction({
        TypeObjet: "Type de tâche administrative",
        ObjetId: createdId,
        Action: ActionHistorique.Creation,
        NouvelleValeur: JSON.stringify(payload)
      });
    }
  }
 
  public async deleteTypeTacheAdministrative(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.TypesTachesAdministratives).items.getById(id).delete();
    await this.safeTracerAction({
      TypeObjet: "Type de tâche administrative",
      ObjetId: id,
      Action: ActionHistorique.Modification,
      Commentaire: "Suppression du type de tâche administrative"
    });
  }

  // --- Périodes de feuille de temps ---

  public async getPeriodes(): Promise<ITimesheetPeriod[]> {
    const sp = getSP(this.context);
    const results: ITimesheetPeriod[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.PeriodesFeuilleTemps).items
      .orderBy("DateDebut", false).top(500)) {
      results.push(...(batch as ITimesheetPeriod[]));
    }

    return results;
  }

  private sanitizePeriode(periode: Partial<ITimesheetPeriod>): Record<string, unknown> {
    const allowedFields: Array<keyof ITimesheetPeriod> = [
      "Title", "CodePeriode", "TypePeriode", "DateDebut", "DateFin",
      "DateLimiteSoumission", "StatutPeriode", "Annee", "NumeroSemaine", "Mois", "Actif"
    ];
    const payload: Record<string, unknown> = {};
    for (const key of allowedFields) {
      const value = periode[key];
      if (value !== undefined && value !== null) { payload[key] = value; }
    }
    return payload;
  }

  public async savePeriode(periode: Partial<ITimesheetPeriod>): Promise<void> {
    const sp = getSP(this.context);
    const payload = this.sanitizePeriode(periode);
    if (periode.Id) {
      await sp.web.lists.getByTitle(LISTS.PeriodesFeuilleTemps).items.getById(periode.Id).update(payload);
      await this.safeTracerAction({
        TypeObjet: "Période de feuille de temps",
        ObjetId: periode.Id,
        Action: ActionHistorique.Modification,
        NouvelleValeur: JSON.stringify(payload)
      });
    } else {
      const result = await sp.web.lists.getByTitle(LISTS.PeriodesFeuilleTemps).items.add(payload);
      const createdId = this.extractAddedItemId(result);
      await this.safeTracerAction({
        TypeObjet: "Période de feuille de temps",
        ObjetId: createdId,
        Action: ActionHistorique.Creation,
        NouvelleValeur: JSON.stringify(payload)
      });
    }
  }

  // Flux 1 - Génère automatiquement les semaines de feuille de temps sur une plage de dates
  public genererPeriodesHebdomadaires(
    anneeDebut: number,
    semaineDebut: number,
    nombreSemaines: number,
    dateDebut: string,
    statutPeriode: StatutPeriode
  ): Partial<ITimesheetPeriod>[] {
    const periodes: Partial<ITimesheetPeriod>[] = [];
    const startDate = new Date(`${dateDebut}T00:00:00`);
    for (let i = 0; i < nombreSemaines; i++) {
      const dateDebutSemaine = new Date(startDate);
      dateDebutSemaine.setDate(startDate.getDate() + i * 7);
      const dateFinSemaine = new Date(dateDebutSemaine);
      dateFinSemaine.setDate(dateDebutSemaine.getDate() + 6);
      const numeroSemaine = ((semaineDebut - 1 + i) % 52) + 1;
      const annee = anneeDebut + Math.floor((semaineDebut - 1 + i) / 52);
      periodes.push({
        Title: `Semaine ${numeroSemaine}-${annee}`,
        CodePeriode: `S${numeroSemaine}-${annee}`,
        TypePeriode: "Hebdomadaire",
        StatutPeriode: statutPeriode,
        DateDebut: dateDebutSemaine.toISOString(),
        DateFin: dateFinSemaine.toISOString(),
        Annee: annee,
        NumeroSemaine: numeroSemaine,
        Actif: true
      });
    }
    return periodes;
  }

  // --- Paramètres application ---

  public async getParametres(): Promise<IApplicationSetting[]> {
    const sp = getSP(this.context);
    const results: IApplicationSetting[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.ParametresApplication).items
      .filter("Actif eq 1").top(500)) {
      results.push(...(batch as IApplicationSetting[]));
    }

    return results;
  }

  private sanitizeApplicationSetting(param: Partial<IApplicationSetting>): Record<string, unknown> {
    const allowedFields: Array<keyof IApplicationSetting> = ["Cle", "Valeur", "TypeValeur", "Description", "Actif"];
    const payload: Record<string, unknown> = {};
    for (const key of allowedFields) {
      const value = param[key];
      if (value === undefined || value === null) { continue; }
      payload[key] = value;
    }
    return payload;
  }

  public async saveParametre(param: Partial<IApplicationSetting>): Promise<void> {
    const sp = getSP(this.context);
    const payload = this.sanitizeApplicationSetting(param);
    if (param.Id) {
      await sp.web.lists.getByTitle(LISTS.ParametresApplication).items.getById(param.Id).update(payload);
      await this.safeTracerAction({
        TypeObjet: "Paramètre application",
        ObjetId: param.Id,
        Action: ActionHistorique.Modification,
        NouvelleValeur: JSON.stringify(payload)
      });
    } else {
      const result = await sp.web.lists.getByTitle(LISTS.ParametresApplication).items.add(payload);
      const createdId = this.extractAddedItemId(result);
      await this.safeTracerAction({
        TypeObjet: "Paramètre application",
        ObjetId: createdId,
        Action: ActionHistorique.Creation,
        NouvelleValeur: JSON.stringify(payload)
      });
    }
  }

  public async deleteParametre(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.ParametresApplication).items.getById(id).delete();
    await this.safeTracerAction({
      TypeObjet: "Paramètre application",
      ObjetId: id,
      Action: ActionHistorique.Modification,
      Commentaire: "Suppression du paramètre application"
    });
  }

  // --- Historique / audit ---

  // Trace une action dans HistoriqueActions ; renseigne automatiquement l'utilisateur connecté
  // si UtilisateurId n'est pas fourni explicitement dans l'entrée.
  public async tracerAction(entree: Partial<IAuditHistory>): Promise<void> {
    const sp = getSP(this.context);
    let utilisateurId = entree.UtilisateurId;
    if (!utilisateurId) {
      const utilisateur = await this.getUtilisateurConnecte();
      utilisateurId = utilisateur.id;
    }
    await sp.web.lists.getByTitle(LISTS.HistoriqueActions).items.add({
      ...entree,
      UtilisateurId: utilisateurId,
      DateAction: new Date().toISOString()
    });
  }

  private async safeTracerAction(entree: Partial<IAuditHistory>): Promise<void> {
    try {
      await this.tracerAction(entree);
    } catch {
      // L'historisation ne doit jamais bloquer l'action métier principale.
    }
  }

  public async getHistorique(typeObjet: string, objetId: number): Promise<IAuditHistory[]> {
    const sp = getSP(this.context);
    const results: IAuditHistory[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.HistoriqueActions).items
      .filter(`TypeObjet eq '${typeObjet}' and ObjetId eq ${objetId}`)
      .orderBy("DateAction", false).top(500)) {
      results.push(...(batch as IAuditHistory[]));
    }

    return results;
  }
}