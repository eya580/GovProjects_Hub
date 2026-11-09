import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { getSP, LISTS, mapLookupIds, buildLookupWritePayload, filtreLookupEgal } from "./SPConfig";
import { IResource, ICapacity } from "../models/IResource";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// RessourceId est une colonne Lookup (titrée avec le suffixe "Id") sur CapacitesRessources
const CHAMPS_LOOKUP_CAPACITE = ["RessourceId"];

export class ResourceService {
  constructor(private context: WebPartContext) {}

  public async getAll(actifSeulement = true): Promise<IResource[]> {
    const sp = getSP(this.context);
    let query = sp.web.lists.getByTitle(LISTS.Ressources).items
      .select("Id", "Title", "UtilisateurId", "Utilisateur/Title", "Utilisateur/EMail",
        "Email", "Matricule", "Profil", "Departement", "ManagerId", "Manager/Title",
        "CapaciteJournaliereHeures", "CapaciteHebdomadaireHeures", "TauxAffectationMaximum",
        "DateEntree", "DateSortie", "Actif", "PeutSaisirTemps", "PeutGererPlanning")
      .expand("Utilisateur", "Manager")
      .orderBy("Title", true);
    if (actifSeulement) { query = query.filter("Actif eq 1"); }

    const results: IResource[] = [];
    for await (const batch of query.top(500)) {
      results.push(...(batch as IResource[]));
    }
    return results;
  }

  // Identifie la fiche Ressource correspondant à l'utilisateur SharePoint connecté
  public async getRessourceConnectee(): Promise<IResource | undefined> {
    const sp = getSP(this.context);
    const currentUser = await sp.web.currentUser();
    const items = await sp.web.lists.getByTitle(LISTS.Ressources).items
      .filter(`UtilisateurId eq ${currentUser.Id} and Actif eq 1`)
      .top(1)() as IResource[];
    return items[0];
  }

  public async create(ressource: Partial<IResource>): Promise<IResource> {
    const sp = getSP(this.context);
    const result = await sp.web.lists.getByTitle(LISTS.Ressources).items.add(ressource);
    return result as unknown as IResource;
  }

  public async update(id: number, ressource: Partial<IResource>): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.Ressources).items.getById(id).update(ressource);
  }

  public async getCapacites(ressourceId: number, dateDebut: string, dateFin: string): Promise<ICapacity[]> {
    try {
      const sp = getSP(this.context);
      const results: ICapacity[] = [];

      for await (const batch of sp.web.lists.getByTitle(LISTS.CapacitesRessources).items
        .select("Id", "RessourceId/Id", "DateDebut", "DateFin", "CapaciteHeuresParJour", "TauxDisponibilite", "CapaciteTotaleHeures", "Motif", "Actif")
        .expand("RessourceId")
        .filter(`${filtreLookupEgal("RessourceId", ressourceId)} and Actif eq 1 and DateFin ge datetime'${dateDebut}T00:00:00Z' and DateDebut le datetime'${dateFin}T23:59:59Z'`)
        .top(500)) {
        results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_CAPACITE) as unknown as ICapacity));
      }

      return results;
    } catch {
      // Liste CapacitesRessources absente ou inaccessible → aucune indisponibilité planifiée
      return [];
    }
  }

  // Retourne toutes les indisponibilités d'une ressource (sans filtre de date — pour la gestion)
  public async getAllCapacites(ressourceId: number): Promise<ICapacity[]> {
    try {
      const sp = getSP(this.context);
      const results: ICapacity[] = [];
      for await (const batch of sp.web.lists.getByTitle(LISTS.CapacitesRessources).items
        .select("Id", "RessourceId/Id", "DateDebut", "DateFin", "CapaciteHeuresParJour", "TauxDisponibilite", "CapaciteTotaleHeures", "Motif", "Actif")
        .expand("RessourceId")
        .filter(filtreLookupEgal("RessourceId", ressourceId))
        .orderBy("DateDebut", false)
        .top(500)) {
        results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_CAPACITE) as unknown as ICapacity));
      }
      return results;
    } catch {
      return [];
    }
  }

  public async createCapacite(capacite: Partial<ICapacity>): Promise<void> {
    const sp = getSP(this.context);
    const payload = buildLookupWritePayload(capacite as Record<string, unknown>, CHAMPS_LOOKUP_CAPACITE);
    await sp.web.lists.getByTitle(LISTS.CapacitesRessources).items.add(payload);
  }

  public async updateCapacite(id: number, capacite: Partial<ICapacity>): Promise<void> {
    const sp = getSP(this.context);
    const allowedFields: Array<keyof ICapacity> = [
      "RessourceId", "DateDebut", "DateFin", "CapaciteHeuresParJour",
      "TauxDisponibilite", "CapaciteTotaleHeures", "Motif", "Actif"
    ];
    const sanitized: Record<string, unknown> = {};
    for (const key of allowedFields) {
      if (capacite[key] !== undefined && capacite[key] !== null) { sanitized[key] = capacite[key]; }
    }
    const payload = buildLookupWritePayload(sanitized, CHAMPS_LOOKUP_CAPACITE);
    await sp.web.lists.getByTitle(LISTS.CapacitesRessources).items.getById(id).update(payload);
  }

  public async deleteCapacite(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.CapacitesRessources).items.getById(id).delete();
  }

  // Capacité nette = capacité théorique - congés/indisponibilités sur la période
  public calculerCapaciteNette(capaciteTheorique: number, heuresAbsence: number): number {
    return Math.max(0, capaciteTheorique - heuresAbsence);
  }

  // Taux de charge = charge affectée / capacité nette x 100
  public calculerTauxDeCharge(chargeAffectee: number, capaciteNette: number): number {
    if (capaciteNette === 0) { return 0; }
    return Math.round((chargeAffectee / capaciteNette) * 10000) / 100;
  }
}