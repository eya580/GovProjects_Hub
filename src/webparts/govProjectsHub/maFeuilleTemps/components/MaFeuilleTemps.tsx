import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  Stack, Spinner, MessageBar, MessageBarType, Dropdown, IDropdownOption,
  DetailsList, IColumn, SelectionMode, PrimaryButton, DefaultButton, TextField, IconButton
} from "@fluentui/react";
import { IMaFeuilleTempsProps } from "./IMaFeuilleTempsProps";
import { ResourceService } from "../../../../services/ResourceService";
import { AssignmentService } from "../../../../services/AssignmentService";
import { TimesheetService } from "../../../../services/TimesheetService";
import { AdminService } from "../../../../services/AdminService";
import { ActivityService } from "../../../../services/ActivityService";
import { ProjectService } from "../../../../services/ProjectService";
import { IResource } from "../../../../models/IResource";
import { IAssignment } from "../../../../models/IAssignment";
import { IActivity, IActivityPredecessor } from "../../../../models/IActivity";
import { ITimesheet, ITimesheetLine, ITimesheetPeriod, ITimesheetRow, IAdminTaskType } from "../../../../models/ITimesheet";
import { TypeLigne, StatutFeuille, StatutLigne, SourceSaisie } from "../../../../models/Enums";
import BiatHeader, {
  biatContainerStyle, biatCardStyle, biatToolbarStyle,
  biatPrimaryBtn, biatDefaultBtn, BIAT_NAVY, BIAT_BORDER
} from "../../components/BiatHeader";

const formatDateKey = (date: Date): string => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const parseDateKey = (value: string): Date => {
  if (value.includes("T")) {
    const d = new Date(value);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
};

const genererDates = (debut: string, fin: string): string[] => {
  const dates: string[] = [];
  let current = parseDateKey(debut);
  const end = parseDateKey(fin);

  while (current.getTime() <= end.getTime()) {
    dates.push(formatDateKey(current));
    current = new Date(current.getFullYear(), current.getMonth(), current.getDate() + 1);
  }
  return dates;
};

const MaFeuilleTemps: React.FC<IMaFeuilleTempsProps> = (props) => {
  const resourceService = useMemo(() => new ResourceService(props.context), [props.context]);
  const assignmentService = useMemo(() => new AssignmentService(props.context), [props.context]);
  const timesheetService = useMemo(() => new TimesheetService(props.context), [props.context]);
  const adminService = useMemo(() => new AdminService(props.context), [props.context]);
  const activityService = useMemo(() => new ActivityService(props.context), [props.context]);
  const projectService  = useMemo(() => new ProjectService(props.context),  [props.context]);

  const [ressource, setRessource] = useState<IResource>();
  const [periodes, setPeriodes] = useState<ITimesheetPeriod[]>([]);
  const [periodeId, setPeriodeId] = useState<number | undefined>(undefined);
  const [feuille, setFeuille] = useState<ITimesheet>();
  const [lignes, setLignes] = useState<ITimesheetLine[]>([]);
  const [affectations, setAffectations] = useState<IAssignment[]>([]);
  const [typesAdmin, setTypesAdmin] = useState<IAdminTaskType[]>([]);
  const [chargement, setChargement] = useState<boolean>(true);
  const [erreur, setErreur] = useState<string>("");
  const debounceRefs = useRef<{ [key: string]: ReturnType<typeof setTimeout> }>({});

  // ── Prédécesseurs ────────────────────────────────────────────────────────────
  const [projetIdPred, setProjetIdPred]         = useState<number | undefined>(undefined);
  const [predecesseurs, setPredecesseurs]         = useState<IActivityPredecessor[]>([]);
  const [chargementPred, setChargementPred]       = useState<boolean>(false);
  const [erreurPred, setErreurPred]               = useState<string>("");
  const [newActiviteId, setNewActiviteId]         = useState<number | undefined>(undefined);
  const [newPredecesseurId, setNewPredecesseurId] = useState<number | undefined>(undefined);
  const [newTypeLien, setNewTypeLien]             = useState<string>("FS");
  const [newDecalage, setNewDecalage]             = useState<string>("0");
  // Données enrichies depuis listes SharePoint
  const [projetsMap, setProjetsMap]               = useState<Map<number, string>>(new Map());
  const [activitesPred, setActivitesPred]         = useState<IActivity[]>([]);
  const [toutesAffectations, setToutesAffectations] = useState<IAssignment[]>([]);
  const [activitesMap, setActivitesMap]           = useState<Map<number, string>>(new Map());

  useEffect(() => {
    (async () => {
      try {
        const [r, p, types] = await Promise.all([
          resourceService.getRessourceConnectee(),
          timesheetService.getAllPeriodes(),
          adminService.getTypesTachesAdministratives()
        ]);
        console.log("Périodes ouvertes reçues", p);
        setRessource(r);
        setPeriodes(p);
        setTypesAdmin(types);
        if (p.length > 0) { setPeriodeId(p[0].Id !== undefined ? Number(p[0].Id) : undefined); }
      } catch (e) {
        console.error("Erreur initialisation MaFeuilleTemps", e);
        setErreur("Erreur d'initialisation : " + (e as Error).message);
      } finally {
        setChargement(false);
      }
    })().catch(() => { /* géré */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const chargerFeuille = async (): Promise<void> => {
    if (!ressource || !periodeId) { return; }
    setChargement(true);
    setErreur("");
    try {
      const utilisateurId = (props.context.pageContext.legacyPageContext as { userId?: number })?.userId || 0;
      const periodeIdNumber = Number(periodeId);
      const f = await timesheetService.getOrCreateFeuille(ressource.Id as number, periodeIdNumber, utilisateurId);
      const [l, affs] = await Promise.all([
        timesheetService.getLignes(f.Id as number),
        assignmentService.getByRessource(ressource.Id as number)
      ]);
      setFeuille(f);
      setLignes(l);
      setAffectations(affs);
    } catch (e) {
      setErreur("Erreur lors du chargement de la feuille : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    chargerFeuille().catch(() => { /* géré */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ressource, periodeId]);

  // Charge tous les projets + toutes les affectations de la ressource
  useEffect(() => {
    if (!ressource) { return; }
    (async () => {
      try {
        const [projets, toutesAffs] = await Promise.all([
          projectService.getAll(),
          assignmentService.getAllByRessource(ressource.Id as number),
        ]);
        const map = new Map<number, string>();
        projets.forEach(p => { if (p.Id !== undefined) { map.set(p.Id as number, p.Title); } });
        setProjetsMap(map);
        setToutesAffectations(toutesAffs);

        // Charger les noms des activités pour toutes les affectations
        const projetIds = [...new Set(toutesAffs.map(a => a.ProjetId))];
        const allActivites = await Promise.all(projetIds.map(id => activityService.getByProjet(id)));
        const actMap = new Map<number, string>();
        ([] as IActivity[]).concat(...allActivites).forEach((act: IActivity) => { if (act.Id !== undefined) { actMap.set(act.Id as number, act.Title); } });
        setActivitesMap(actMap);
      } catch { /* silencieux */ }
    })().catch(() => { /* géré */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ressource]);

  const periode = periodes.filter(p => p.Id === periodeId)[0];
  const dates = periode ? genererDates(periode.DateDebut, periode.DateFin) : [];

  const periodeOptions = useMemo(() => periodes
    .filter(p => p.Id !== undefined && p.Id !== null)
    .map(p => ({ key: Number(p.Id), text: p.Title || `Période ${p.Id}` } as IDropdownOption)),
  [periodes]);

  const lignesGroupees: ITimesheetRow[] = useMemo(() => {
    const map: { [cle: string]: ITimesheetRow } = {};

    affectations.filter(a => a.ActiviteId !== 4).forEach(a => {
      const cle = `AFF-${a.Id}`;
      map[cle] = { cle, libelle: activitesMap.get(a.ActiviteId) || `Tâche #${a.ActiviteId}`, type: TypeLigne.Projet, affectationId: a.Id, jours: {}, totalLigne: 0 };
    });
    typesAdmin.forEach(t => {
      const cle = `ADM-${t.Id}`;
      map[cle] = { cle, libelle: t.Title, type: TypeLigne.Administrative, typeTacheAdministrativeId: t.Id, jours: {}, totalLigne: 0 };
    });

    lignes.forEach(l => {
      const cle = l.AffectationId ? `AFF-${l.AffectationId}` : `ADM-${l.TypeTacheAdministrativeId}`;
      if (!map[cle]) { return; }
      const dateKey = formatDateKey(parseDateKey(l.DateTravail));
      const existant = map[cle].jours[dateKey];
      if (!existant || (l.Id !== undefined && (existant.ligne?.Id === undefined || l.Id > existant.ligne.Id))) {
        map[cle].jours[dateKey] = { date: l.DateTravail, ligne: l, heures: l.NombreHeures };
      }
    });

    Object.values(map).forEach(row => {
      row.totalLigne = Object.values(row.jours).reduce((s, c) => s + (c.heures || 0), 0);
    });

    return Object.values(map);
  }, [affectations, typesAdmin, lignes, activitesMap]);

  const modifiable = feuille ? !feuille.Verrouillee && feuille.StatutFeuille !== StatutFeuille.Validee : false;

  const saisirHeures = (row: ITimesheetRow, date: string, heures: number): void => {
    if (!feuille || !ressource) { return; }
    const ligneExistante = row.jours[date]?.ligne;
    const payload: Partial<ITimesheetLine> = {
      Id: ligneExistante?.Id,
      FeuilleTempsId: feuille.Id,
      RessourceId: ressource.Id,
      DateTravail: date,
      TypeLigne: row.type,
      AffectationId: row.affectationId,
      TypeTacheAdministrativeId: row.typeTacheAdministrativeId,
      NombreHeures: heures,
      StatutLigne: StatutLigne.Brouillon,
      SourceSaisie: SourceSaisie.Web,
      Modifiable: true,
      DateCreationSaisie: new Date().toISOString()
    };

    setLignes(prev => {
      const index = prev.findIndex(l =>
        (ligneExistante?.Id !== undefined && l.Id === ligneExistante.Id) || (
          l.FeuilleTempsId === feuille.Id &&
          formatDateKey(parseDateKey(l.DateTravail)) === date &&
          l.AffectationId === row.affectationId &&
          l.TypeTacheAdministrativeId === row.typeTacheAdministrativeId
        )
      );

      if (index >= 0) {
        const updated = [...prev];
        updated[index] = { ...updated[index], ...payload, DateTravail: date, NombreHeures: heures } as ITimesheetLine;
        return updated;
      }

      return [...prev, {
        ...payload,
        FeuilleTempsId: feuille.Id as number,
        RessourceId: ressource.Id as number,
        DateTravail: date,
        NombreHeures: heures,
        TypeLigne: row.type,
        AffectationId: row.affectationId,
        TypeTacheAdministrativeId: row.typeTacheAdministrativeId,
        StatutLigne: StatutLigne.Brouillon,
        SourceSaisie: SourceSaisie.Web,
        Modifiable: true,
        DateCreationSaisie: new Date().toISOString()
      } as ITimesheetLine];
    });

    const debounceKey = `${row.cle}-${date}`;
    clearTimeout(debounceRefs.current[debounceKey]);
    debounceRefs.current[debounceKey] = setTimeout(() => {
      timesheetService.upsertLigne(payload)
        .then(() => chargerFeuille())
        .catch(e => {
          setErreur("Erreur lors de la mise à jour de la saisie : " + (e as Error).message);
        });
    }, 600);
  };

  const soumettre = async (): Promise<void> => {
    if (!feuille) { return; }
    try {
      await timesheetService.soumettre(feuille, lignes);
      await chargerFeuille();
    } catch (e) {
      setErreur((e as Error).message);
    }
  };

  // ── Logique prédécesseurs ─────────────────────────────────────────────────────
  const projetsUniques: IDropdownOption[] = useMemo(() =>
    Array.from(projetsMap.entries()).map(([id, title]) => ({ key: id, text: title } as IDropdownOption)),
  [projetsMap]);

  const activiteOptions: IDropdownOption[] = useMemo(() =>
    activitesPred.map(a => ({ key: a.Id as number, text: `${a.WBS ? a.WBS + ' – ' : ''}${a.Title}` } as IDropdownOption)),
  [activitesPred]);

  const activitesPredMap = useMemo(() => {
    const m = new Map<number, string>();
    activitesPred.forEach(a => { if (a.Id !== undefined) { m.set(a.Id as number, a.Title); } });
    return m;
  }, [activitesPred]);

  const chargerPredecesseurs = async (projetId: number): Promise<void> => {
    setChargementPred(true);
    setErreurPred("");
    setNewActiviteId(undefined);
    setNewPredecesseurId(undefined);
    try {
      const [data, toutesActivites] = await Promise.all([
        activityService.getPredecesseurs(projetId),
        activityService.getByProjet(projetId),
      ]);
      setPredecesseurs(data);
      setActivitesPred(toutesActivites);
    } catch (e) {
      setErreurPred("Erreur chargement prédécesseurs : " + (e as Error).message);
    } finally {
      setChargementPred(false);
    }
  };

  const ajouterPredecesseur = async (): Promise<void> => {
    if (!projetIdPred || !newActiviteId || !newPredecesseurId) {
      setErreurPred("Veuillez sélectionner un projet, une activité et un prédécesseur.");
      return;
    }
    setChargementPred(true);
    setErreurPred("");
    try {
      await activityService.addPredecesseur({
        ProjetId:       projetIdPred,
        ActiviteId:     newActiviteId,
        PredecesseurId: newPredecesseurId,
        TypeLien:       newTypeLien,
        DecalageJours:  Number(newDecalage),
        Actif:          true
      });
      setNewActiviteId(undefined);
      setNewPredecesseurId(undefined);
      setNewTypeLien("FS");
      setNewDecalage("0");
      await chargerPredecesseurs(projetIdPred);
    } catch (e) {
      setErreurPred("Erreur ajout prédécesseur : " + (e as Error).message);
    } finally {
      setChargementPred(false);
    }
  };

  const supprimerPredecesseur = async (id: number | undefined): Promise<void> => {
    if (id === undefined || !projetIdPred) { return; }
    setChargementPred(true);
    setErreurPred("");
    try {
      await activityService.deletePredecesseur(id);
      await chargerPredecesseurs(projetIdPred);
    } catch (e) {
      setErreurPred("Erreur suppression prédécesseur : " + (e as Error).message);
    } finally {
      setChargementPred(false);
    }
  };

  const columnsPred: IColumn[] = [
    { key: "act",  name: "Activité",     fieldName: "ActiviteId",     minWidth: 140, maxWidth: 220,
      onRender: (r: IActivityPredecessor) => <span>{activitesPredMap.get(r.ActiviteId) || `#${r.ActiviteId}`}</span> },
    { key: "pred", name: "Prédécesseur", fieldName: "PredecesseurId", minWidth: 140, maxWidth: 220,
      onRender: (r: IActivityPredecessor) => <span>{activitesPredMap.get(r.PredecesseurId) || `#${r.PredecesseurId}`}</span> },
    { key: "type", name: "Type lien",    fieldName: "TypeLien",        minWidth: 80,  maxWidth: 100 },
    { key: "dec",  name: "Décalage (j)", fieldName: "DecalageJours",   minWidth: 90,  maxWidth: 110,
      onRender: (r: IActivityPredecessor) => <span>{r.DecalageJours ?? 0}</span> },
    { key: "del",  name: "",             minWidth: 40,  maxWidth: 50,
      onRender: (r: IActivityPredecessor) => (
        <IconButton
          iconProps={{ iconName: "Delete" }}
          title="Supprimer"
          styles={{ root: { color: "#C50F1F" } }}
          onClick={() => supprimerPredecesseur(r.Id).catch(() => { /* géré */ })}
        />
      )},
  ];

  const columnsAff: IColumn[] = [
    { key: "proj",  name: "Projet",  minWidth: 160, maxWidth: 220,
      onRender: (a: IAssignment) => <span>{projetsMap.get(a.ProjetId) || `Projet #${a.ProjetId}`}</span> },
    { key: "act",   name: "Tâche",   minWidth: 140, maxWidth: 220,
      onRender: (a: IAssignment) => <span>{activitesMap.get(a.ActiviteId) || `#${a.ActiviteId}`}</span> },
    { key: "debut", name: "Début",   fieldName: "DateDebutAffectation", minWidth: 90, maxWidth: 110 },
    { key: "fin",   name: "Fin",     fieldName: "DateFinAffectation",   minWidth: 90, maxWidth: 110 },
    { key: "role",  name: "Rôle",    fieldName: "RoleProjet",           minWidth: 120, maxWidth: 180 },
    { key: "resp",  name: "Responsable", minWidth: 100, maxWidth: 120,
      onRender: (a: IAssignment) => <span>{a.ResponsableTache ? "Oui" : "Non"}</span> },
  ];

  const sectionStyle: React.CSSProperties = {
    marginTop: 28,
    borderTop: `2px solid #E8EEF6`,
    paddingTop: 20,
  };

  const sectionTitleStyle: React.CSSProperties = {
    fontSize: 15,
    fontWeight: 700,
    color: BIAT_NAVY,
    marginBottom: 14,
    fontFamily: "'Segoe UI', sans-serif",
  };

  const columns: IColumn[] = [
    { key: "libelle", name: "Tâche / Activité administrative", fieldName: "libelle", minWidth: 200, maxWidth: 260 },
    ...dates.map(d => ({
      key: d,
      name: parseDateKey(d).toLocaleDateString(undefined, { weekday: "short", day: "2-digit", month: "2-digit" }),
      minWidth: 70, maxWidth: 80,
      onRender: (row: ITimesheetRow) => (
        <TextField
          type="number"
          value={String(row.jours[d]?.heures ?? "")}
          disabled={!modifiable}
          onChange={(_, v) => saisirHeures(row, d, Number(v || 0))}
        />
      )
    })),
    {
      key: "total", name: "Total", minWidth: 60, maxWidth: 70,
      onRender: (row: ITimesheetRow) => (
        <span style={{ fontWeight: '600', color: BIAT_NAVY }}>{row.totalLigne}</span>
      )
    }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title="Ma feuille de temps" subtitle="Saisie et soumission de vos heures de travail" />
      <div style={biatCardStyle}>

        {/* ── Feuille de temps ─────────────────────────────────────────────── */}
        {erreur && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 16 } }}>{erreur}</MessageBar>}

        <div style={biatToolbarStyle}>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" wrap>
            <Dropdown
              label="Période"
              options={periodeOptions}
              selectedKey={periodeId}
              onChange={(_, opt) => setPeriodeId(opt?.key !== undefined ? Number(opt.key) : undefined)}
              styles={{ root: { width: 260 } }}
            />
            {feuille && feuille.StatutFeuille !== StatutFeuille.Brouillon && (
              <div style={{ padding: '6px 14px', background: '#EEF2F8', border: `1px solid ${BIAT_BORDER}`, fontSize: 13, color: BIAT_NAVY, alignSelf: 'flex-end', marginBottom: 2 }}>
                Statut : <strong>{feuille.StatutFeuille}</strong>
                {feuille.CommentaireValidateur && ` — Motif rejet : ${feuille.CommentaireValidateur}`}
              </div>
            )}
          </Stack>
        </div>

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}><Spinner label="Chargement..." /></div>
        ) : (
          <>
            <DetailsList items={lignesGroupees} columns={columns} selectionMode={SelectionMode.none} setKey="cle" />
            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
              <PrimaryButton text="Soumettre la feuille" styles={biatPrimaryBtn} onClick={soumettre} disabled={!modifiable} />
              <DefaultButton text="Actualiser" styles={biatDefaultBtn} onClick={() => chargerFeuille().catch(() => { /* géré */ })} />
            </Stack>
          </>
        )}

        {/* ── Mes affectations ─────────────────────────────────────────────── */}
        <div style={sectionStyle}>
          <div style={sectionTitleStyle}>Mes affectations</div>
          {toutesAffectations.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>Aucune affectation trouvée pour votre profil.</MessageBar>
          ) : (
            <DetailsList
              items={toutesAffectations}
              columns={columnsAff}
              selectionMode={SelectionMode.none}
              setKey="aff"
            />
          )}
        </div>

        {/* ── Prédécesseurs d'activité ─────────────────────────────────────── */}
        <div style={sectionStyle}>
          <div style={sectionTitleStyle}>Prédécesseurs d'activité</div>

          {erreurPred && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 12 } }}>{erreurPred}</MessageBar>}

          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end" wrap styles={{ root: { marginBottom: 16 } }}>
            <Dropdown
              label="Projet"
              placeholder="Sélectionner un projet"
              options={projetsUniques}
              selectedKey={projetIdPred ?? null}
              onChange={(_, opt) => {
                const id = opt?.key as number;
                setProjetIdPred(id);
                chargerPredecesseurs(id).catch(() => { /* géré */ });
              }}
              styles={{ root: { width: 200 } }}
            />
          </Stack>

          {chargementPred ? (
            <Spinner label="Chargement des prédécesseurs..." />
          ) : projetIdPred && (
            <>
              <DetailsList
                items={predecesseurs}
                columns={columnsPred}
                selectionMode={SelectionMode.none}
                setKey="pred"
              />

              {/* Formulaire d'ajout */}
              <div style={{ marginTop: 16, padding: '14px 16px', background: '#F2F5F9', border: `1px solid ${BIAT_BORDER}` }}>
                <div style={{ ...sectionTitleStyle, fontSize: 13, marginBottom: 10 }}>Ajouter un lien prédécesseur</div>
                <Stack horizontal tokens={{ childrenGap: 10 }} wrap verticalAlign="end">
                  <Dropdown
                    label="Activité"
                    placeholder="Sélectionner une activité"
                    options={activiteOptions}
                    selectedKey={newActiviteId ?? null}
                    onChange={(_, opt) => setNewActiviteId(opt?.key as number)}
                    styles={{ root: { width: 220 } }}
                  />
                  <Dropdown
                    label="Prédécesseur"
                    placeholder="Sélectionner le prédécesseur"
                    options={activiteOptions}
                    selectedKey={newPredecesseurId ?? null}
                    onChange={(_, opt) => setNewPredecesseurId(opt?.key as number)}
                    styles={{ root: { width: 220 } }}
                  />
                  <Dropdown
                    label="Type de lien"
                    selectedKey={newTypeLien}
                    options={[
                      { key: "FS", text: "FS – Fin → Début" },
                      { key: "SS", text: "SS – Début → Début" },
                      { key: "FF", text: "FF – Fin → Fin" },
                      { key: "SF", text: "SF – Début → Fin" },
                    ]}
                    onChange={(_, opt) => setNewTypeLien(opt?.key as string)}
                    styles={{ root: { width: 180 } }}
                  />
                  <TextField
                    label="Décalage (jours)"
                    type="number"
                    value={newDecalage}
                    onChange={(_, v) => setNewDecalage(v || "0")}
                    styles={{ root: { width: 130 } }}
                  />
                  <PrimaryButton
                    text="Ajouter"
                    styles={biatPrimaryBtn}
                    onClick={ajouterPredecesseur}
                    disabled={chargementPred}
                    style={{ alignSelf: 'flex-end' }}
                  />
                </Stack>
              </div>
            </>
          )}
        </div>

      </div>
    </div>
  );
};

export default MaFeuilleTemps;
