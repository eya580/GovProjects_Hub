import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP, LISTS, mapLookupIds, buildLookupWritePayload, filtreLookupEgal } from "./SPConfig";
import { IAssignment } from "../models/IAssignment";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { StatutAffectation } from "../models/Enums";

// ProjetId, ActiviteId et RessourceId sont des colonnes Lookup (titrées avec le suffixe "Id")
const CHAMPS_LOOKUP_AFFECTATION = ["ProjetId", "ActiviteId", "RessourceId"];

const SELECT_FIELDS = [
  "Id", "ProjetId/Id", "ActiviteId/Id", "RessourceId/Id",
  "DateDebutAffectation", "DateFinAffectation", "ChargeAffecteeHeures", "UniteAffectation",
  "HeuresParJour", "RoleProjet", "ResponsableTache", "AutoriseSaisieTemps", "StatutAffectation", "Actif"
];

export class AssignmentService {
  constructor(private context: WebPartContext) {}

  public async getByProjet(projetId: number): Promise<IAssignment[]> {
    const sp = getSP(this.context);
    const results: IAssignment[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.Affectations).items
      .select(...SELECT_FIELDS)
      .expand("ProjetId", "ActiviteId", "RessourceId")
      .filter(`${filtreLookupEgal("ProjetId", projetId)} and Actif eq 1`)
      .top(500)) {
      results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_AFFECTATION) as unknown as IAssignment));
    }

    return results;
  }

  public async getByRessource(ressourceId: number, dateDebut?: string, dateFin?: string): Promise<IAssignment[]> {
    try {
      const sp = getSP(this.context);
      let filter = `${filtreLookupEgal("RessourceId", ressourceId)} and Actif eq 1 and StatutAffectation eq '${StatutAffectation.Active}'`;
      if (dateDebut && dateFin) {
        filter += ` and DateDebutAffectation le datetime'${dateFin}T23:59:59Z' and DateFinAffectation ge datetime'${dateDebut}T00:00:00Z'`;
      }

      const results: IAssignment[] = [];
      for await (const batch of sp.web.lists.getByTitle(LISTS.Affectations).items
        .select(...SELECT_FIELDS)
        .expand("ProjetId", "ActiviteId", "RessourceId")
        .filter(filter)
        .top(500)) {
        results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_AFFECTATION) as unknown as IAssignment));
      }

      return results;
    } catch {
      // Liste Affectations absente ou inaccessible → aucune charge affectée
      return [];
    }
  }

  public async getAllByRessource(ressourceId: number): Promise<IAssignment[]> {
    const sp = getSP(this.context);
    const results: IAssignment[] = [];
    for await (const batch of sp.web.lists.getByTitle(LISTS.Affectations).items
      .select(...SELECT_FIELDS)
      .expand("ProjetId", "ActiviteId", "RessourceId")
      .filter(`${filtreLookupEgal("RessourceId", ressourceId)} and Actif eq 1`)
      .top(500)) {
      results.push(...(batch as Record<string, unknown>[]).map(i => mapLookupIds(i, CHAMPS_LOOKUP_AFFECTATION) as unknown as IAssignment));
    }
    return results;
  }

  public async create(affectation: Partial<IAssignment>): Promise<IAssignment> {
    const sp = getSP(this.context);
    const payload = buildLookupWritePayload(affectation as Record<string, unknown>, CHAMPS_LOOKUP_AFFECTATION);
    const result = await sp.web.lists.getByTitle(LISTS.Affectations).items.add(payload);
    return result as unknown as IAssignment;
  }

  public async update(id: number, affectation: Partial<IAssignment>): Promise<void> {
    const sp = getSP(this.context);
    const payload = buildLookupWritePayload(affectation as Record<string, unknown>, CHAMPS_LOOKUP_AFFECTATION);
    await sp.web.lists.getByTitle(LISTS.Affectations).items.getById(id).update(payload);
  }

  public async delete(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.Affectations).items.getById(id).delete();
  }

  // Une ressource ne peut saisir du temps que sur une affectation active couvrant la date saisie
  public estAffectationValidePourSaisie(affectation: IAssignment, dateTravail: string): boolean {
    return affectation.Actif === true
      && affectation.AutoriseSaisieTemps === true
      && affectation.StatutAffectation === StatutAffectation.Active
      && dateTravail >= affectation.DateDebutAffectation
      && dateTravail <= affectation.DateFinAffectation;
  }

  // Détecte une surcharge : somme des charges affectées > capacité nette disponible
  public detecterSurcharge(chargeTotaleAffecteeHeures: number, capaciteNetteHeures: number): boolean {
    return chargeTotaleAffecteeHeures > capaciteNetteHeures;
  }
}