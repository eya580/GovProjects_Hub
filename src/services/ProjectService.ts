import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getSP, LISTS } from "./SPConfig";
import { IProject } from "../models/IProject";
import { WebPartContext } from "@microsoft/sp-webpart-base";

const SELECT_FIELDS = [
  "Id", "Title", "CodeProjet", "Description", "ChefProjetId", "ChefProjet/Title", "ChefProjet/EMail",
  "SponsorId", "Sponsor/Title", "Sponsor/EMail", "Direction", "StatutProjet",
  "DateDebutReference", "DateFinReference", "DateDebutPrevue", "DateFinPrevue",
  "DateDebutReelle", "DateFinReelle", "PourcentageAchevement", "Priorite", "Budget",
  "Actif", "SiteProjetUrl", "Author/Title", "Editor/Title"
];
const EXPAND_FIELDS = ["ChefProjet", "Sponsor", "Author", "Editor"];

export class ProjectService {
  constructor(private context: WebPartContext) {}

  public async getAll(filtreStatut?: string, filtreChefProjetId?: number): Promise<IProject[]> {
    const sp = getSP(this.context);
    let query = sp.web.lists.getByTitle(LISTS.Projets).items
      .select(...SELECT_FIELDS)
      .expand(...EXPAND_FIELDS)
      .orderBy("Title", true);

    const filters: string[] = [];
    if (filtreStatut) { filters.push(`StatutProjet eq '${filtreStatut}'`); }
    if (filtreChefProjetId) { filters.push(`ChefProjetId eq ${filtreChefProjetId}`); }
    if (filters.length > 0) { query = query.filter(filters.join(" and ")); }

    const results: IProject[] = [];
    for await (const batch of query.top(500)) {
      results.push(...(batch as IProject[]));
    }
    return results;
  }

  public async getById(id: number): Promise<IProject> {
    const sp = getSP(this.context);
    return sp.web.lists.getByTitle(LISTS.Projets).items.getById(id)
      .select(...SELECT_FIELDS).expand(...EXPAND_FIELDS)() as Promise<IProject>;
  }

  public async create(project: Partial<IProject>): Promise<IProject> {
    const sp = getSP(this.context);
    const result = await sp.web.lists.getByTitle(LISTS.Projets).items.add(project);
    return result as unknown as IProject;
  }

  public async update(id: number, project: Partial<IProject>): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.Projets).items.getById(id).update(project);
  }

  public async delete(id: number): Promise<void> {
    const sp = getSP(this.context);
    await sp.web.lists.getByTitle(LISTS.Projets).items.getById(id).delete();
  }

  // Recalcule le % d'avancement d'un projet à partir des tâches (charge planifiée pondérée)
  public calculerAvancementPondere(taches: { pourcentage: number; chargePlanifiee: number }[]): number {
    const totalCharge = taches.reduce((sum, t) => sum + (t.chargePlanifiee || 0), 0);
    if (totalCharge === 0) { return 0; }
    const somme = taches.reduce((sum, t) => sum + t.pourcentage * t.chargePlanifiee, 0);
    return Math.round((somme / totalCharge) * 100) / 100;
  }
}