import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP, LISTS } from "./SPConfig";
import {
  IActivity, IActivityNode, IActivityPredecessor,
  IActivitySP, IActivityPredecessorSP, IActivityWritePayload
} from "../models/IActivity";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { TypeActivite } from "../models/Enums";

const SELECT_FIELDS = [
  "Id", "Title", "CodeActivite", "ProjetId/Id", "TypeActivite", "ParentActiviteId/Id",
  "WBS", "NiveauWBS", "Ordre", "Description", "DateDebutReference", "DateFinReference",
  "DateDebutPrevue", "DateFinPrevue", "DateDebutReelle", "DateFinReelle",
  "PourcentageAchevement", "ChargePlanifieeHeures", "ChargeReelleHeures",
  "DureePlanifieeJours", "StatutActivite", "EstJalon", "AutoriseSaisieTemps",
  "EstActive", "ResponsableActiviteId", "ModeCalculAvancement", "ModifiableParRessource"
];

export class ActivityService {
  constructor(private context: WebPartContext) {}

  public async getByProjet(projetId: number): Promise<IActivity[]> {
    const sp = getSP(this.context);
    const results: IActivity[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.Activites).items
      .select(...SELECT_FIELDS)
      .expand("ProjetId", "ParentActiviteId")
      .filter(`ProjetId/Id eq ${projetId} and EstActive eq 1`)
      .orderBy("Ordre", true)
      .top(500)) {
      results.push(...this.mapLookups(batch as IActivitySP[]));
    }

    return results;
  }

  private mapLookups(items: IActivitySP[]): IActivity[] {
    return items.map(item => ({
      ...item,
      ProjetId: item.ProjetId.Id,
      ParentActiviteId: item.ParentActiviteId?.Id
    }));
  }

  private sanitizeActivite(activite: Partial<IActivity>): Partial<IActivity> {
    const allowedFields: Array<keyof IActivity> = [
      "Title", "CodeActivite", "ProjetId", "TypeActivite", "ParentActiviteId", "WBS", "NiveauWBS", "Ordre",
      "Description", "DateDebutReference", "DateFinReference", "DateDebutPrevue", "DateFinPrevue",
      "DateDebutReelle", "DateFinReelle", "PourcentageAchevement", "ChargePlanifieeHeures",
      "ChargeReelleHeures", "DureePlanifieeJours", "StatutActivite", "EstJalon", "AutoriseSaisieTemps",
      "EstActive", "ResponsableActiviteId", "ModeCalculAvancement", "ModifiableParRessource"
    ];
    const payload = {} as Partial<IActivity>;
    for (const key of allowedFields) {
      const value = activite[key];
      if (value === undefined || value === null) { continue; }
      if (typeof value === "object") { continue; }
      if (typeof value === "number" && Number.isNaN(value)) { continue; }
      // Normalize date-only strings (YYYY-MM-DD) to avoid timezone shifts when stored
      if (typeof value === "string" && key.startsWith("Date") && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
        // Use midday UTC to avoid date shifting across timezones
        (payload as Record<string, unknown>)[key] = `${value}T12:00:00Z`;
      } else {
        (payload as Record<string, unknown>)[key] = value;
      }
    }
    return payload;
  }

  private buildLookupPayload(payload: Partial<IActivity>): IActivityWritePayload {
    const { ProjetId, ParentActiviteId, ...rest } = payload;
    const result: IActivityWritePayload = { ...rest };
    if (ProjetId !== undefined) { result.ProjetIdId = ProjetId; }
    if (ParentActiviteId !== undefined) { result.ParentActiviteIdId = ParentActiviteId; }
    return result;
  }

  public async create(activite: Partial<IActivity>): Promise<IActivity> {
    const sp = getSP(this.context);
    const payload = this.buildLookupPayload(this.sanitizeActivite(activite));
    const result = await sp.web.lists.getByTitle(LISTS.Activites).items.add(payload);
    return result as unknown as IActivity;
  }

  public async update(id: number, activite: Partial<IActivity>): Promise<void> {
    const sp = getSP(this.context);
    const payload = this.buildLookupPayload(this.sanitizeActivite(activite));
    await sp.web.lists.getByTitle(LISTS.Activites).items.getById(id).update(payload);
  }

  public async delete(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.Activites).items.getById(id).delete();
  }

  // Persiste le nouvel ordre (et WBS) d'un lot d'activités après un glisser-déposer.
  // Ces champs ne sont pas des lookups : pas besoin de buildLookupPayload ici.
  public async updateOrdreLot(updates: { id: number; ordre: number; wbs?: string; niveauWBS?: number }[]): Promise<void> {
    const sp = getSP(this.context);
    const list = sp.web.lists.getByTitle(LISTS.Activites);
    for (const u of updates) {
      await list.items.getById(u.id).update({ Ordre: u.ordre, WBS: u.wbs, NiveauWBS: u.niveauWBS });
    }
  }

  public async getPredecesseurs(projetId: number): Promise<IActivityPredecessor[]> {
    const sp = getSP(this.context);
    const results: IActivityPredecessor[] = [];

    for await (const batch of sp.web.lists.getByTitle(LISTS.PredecesseursActivites).items
      .select("Id", "ActiviteId/Id", "PredecesseurId/Id", "TypeLien", "DecalageJours", "ProjetId/Id", "Actif")
      .expand("ProjetId", "ActiviteId", "PredecesseurId")
      .filter(`ProjetId/Id eq ${projetId} and Actif eq 1`)
      .top(500)) {
      results.push(...(batch as IActivityPredecessorSP[]).map(item => ({
        ...item,
        ProjetId:       item.ProjetId.Id,
        ActiviteId:     item.ActiviteId.Id,
        PredecesseurId: item.PredecesseurId.Id,
      })));
    }

    return results;
  }

  public async addPredecesseur(lien: Partial<IActivityPredecessor>): Promise<void> {
    const sp = getSP(this.context);
    const { ProjetId, ActiviteId, PredecesseurId, ...rest } = lien;
    const payload: Record<string, unknown> = { ...rest };
    if (ProjetId       !== undefined) { payload.ProjetIdId       = ProjetId; }
    if (ActiviteId     !== undefined) { payload.ActiviteIdId     = ActiviteId; }
    if (PredecesseurId !== undefined) { payload.PredecesseurIdId = PredecesseurId; }
    await sp.web.lists.getByTitle(LISTS.PredecesseursActivites).items.add(payload);
  }

  public async deletePredecesseur(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.PredecesseursActivites).items.getById(id).delete();
  }

  // --- Logique métier WBS ---

  public buildHierarchy(activites: IActivity[]): IActivityNode[] {
    const map: { [id: number]: IActivityNode } = {};
    activites.forEach(a => { map[a.Id as number] = { ...a, children: [], depth: 0 }; });

    const roots: IActivityNode[] = [];
    activites.forEach(a => {
      const node = map[a.Id as number];
      if (a.ParentActiviteId && map[a.ParentActiviteId]) {
        const parent = map[a.ParentActiviteId];
        node.depth = parent.depth + 1;
        parent.children.push(node);
      } else {
        roots.push(node);
      }
    });

    const sortRec = (nodes: IActivityNode[]): void => {
      nodes.sort((a, b) => (a.Ordre || 0) - (b.Ordre || 0));
      nodes.forEach(n => sortRec(n.children));
    };
    sortRec(roots);
    return roots;
  }

  public recalculerWBS(roots: IActivityNode[]): { id: number; wbs: string; niveau: number }[] {
    const resultats: { id: number; wbs: string; niveau: number }[] = [];
    const walk = (nodes: IActivityNode[], prefix: string, niveau: number): void => {
      nodes.forEach((n, index) => {
        const code = prefix ? `${prefix}.${index + 1}` : `${index + 1}`;
        resultats.push({ id: n.Id as number, wbs: code, niveau });
        walk(n.children, code, niveau + 1);
      });
    };
    walk(roots, "", 1);
    return resultats;
  }

  public calculerAvancementParent(enfants: IActivity[]): number {
    const totalCharge = enfants.reduce((s, a) => s + (a.ChargePlanifieeHeures || 0), 0);
    if (totalCharge === 0) { return 0; }
    const somme = enfants.reduce((s, a) => s + (a.PourcentageAchevement || 0) * (a.ChargePlanifieeHeures || 0), 0);
    return Math.round((somme / totalCharge) * 100) / 100;
  }

  public estSaisissable(activite: IActivity): boolean {
    return activite.TypeActivite === TypeActivite.Tache
      && activite.AutoriseSaisieTemps === true
      && activite.EstActive === true;
  }
}