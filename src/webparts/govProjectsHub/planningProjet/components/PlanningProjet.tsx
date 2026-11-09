import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import {
  DetailsList, IColumn, SelectionMode, Stack, Spinner, MessageBar, MessageBarType,
  TextField, PrimaryButton, DefaultButton, Dropdown, IDropdownOption, Panel, PanelType,
  ProgressIndicator, IDetailsRowProps, IRenderFunction, Icon, IconButton
} from "@fluentui/react";
import { IPlanningProjetProps } from "./IPlanningProjetProps";
import { ActivityService } from "../../../../services/ActivityService";
import { ProjectService } from "../../../../services/ProjectService";
import { AdminService } from "../../../../services/AdminService";
import { IActivity, IActivityNode } from "../../../../models/IActivity";
import { IProject } from "../../../../models/IProject";
import { TypeActivite, StatutActivite, TypeObjetHistorique, ActionHistorique } from "../../../../models/Enums";
import GanttChart from "./GanttChart";
import ComparaisonPlanning from "./ComparaisonPlanning";
import BiatHeader, {
  biatContainerStyle, biatCardStyle, biatToolbarStyle,
  biatPrimaryBtn, biatDefaultBtn, BIAT_NAVY, BIAT_TEAL
} from "../../components/BiatHeader";

const TYPE_OPTIONS: IDropdownOption[] = Object.values(TypeActivite).map(t => ({ key: t, text: t }));
const STATUT_OPTIONS: IDropdownOption[] = Object.values(StatutActivite).map(s => ({ key: s, text: s }));

const INDENT_PAR_NIVEAU = 44;

const PlanningProjet: React.FC<IPlanningProjetProps> = (props) => {
  const activityService = useMemo(() => new ActivityService(props.context), [props.context]);
  const projectService = useMemo(() => new ProjectService(props.context), [props.context]);
  const adminService = useMemo(() => new AdminService(props.context), [props.context]);

  const [projets, setProjets] = useState<IProject[]>([]);
  const [projetId, setProjetId] = useState<number | undefined>(props.projetIdParDefaut);
  const [activites, setActivites] = useState<IActivity[]>([]);
  const [chargement, setChargement] = useState<boolean>(false);
  const [erreur, setErreur] = useState<string>("");
  const [panelOuvert, setPanelOuvert] = useState<boolean>(false);
  const [edition, setEdition] = useState<Partial<IActivity>>({});
  const [ligneGlissee, setLigneGlissee] = useState<IActivityNode | undefined>(undefined);
  const [ganttVisible, setGanttVisible] = useState<boolean>(false);
  const [ganttWidth, setGanttWidth] = useState<number>(480);
  const isDragging = useRef(false);
  const dragStart = useRef({ x: 0, width: 0 });

  useEffect(() => {
    const onMouseMove = (e: MouseEvent): void => {
      if (!isDragging.current) { return; }
      const delta = e.clientX - dragStart.current.x;
      setGanttWidth(() => Math.max(220, Math.min(900, dragStart.current.width + delta)));
    };
    const onMouseUp = (): void => { isDragging.current = false; };
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
    return () => {
      window.removeEventListener("mousemove", onMouseMove);
      window.removeEventListener("mouseup", onMouseUp);
    };
  }, []);

  useEffect(() => {
    projectService.getAll().then(setProjets).catch(() => setErreur("Erreur de chargement des projets."));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const chargerActivites = async (pId: number): Promise<void> => {
    setChargement(true);
    setErreur("");
    try {
      const data = await activityService.getByProjet(pId);
      setActivites(data);
    } catch (e) {
      setErreur("Erreur lors du chargement des activités : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    if (projetId) { chargerActivites(projetId).catch(() => { /* géré */ }); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projetId]);

  const arbre: IActivityNode[] = useMemo(() => activityService.buildHierarchy(activites), [activites]);

  const listeAplatie: IActivityNode[] = useMemo(() => {
    const resultat: IActivityNode[] = [];
    const walk = (nodes: IActivityNode[]): void => {
      nodes.forEach(n => { resultat.push(n); walk(n.children); });
    };
    walk(arbre);
    return resultat;
  }, [arbre]);

  const tracerHistorique = async (
    action: ActionHistorique,
    activiteId: number,
    ancienneValeur?: Record<string, unknown>,
    nouvelleValeur?: Record<string, unknown>,
    commentaire?: string
  ): Promise<void> => {
    try {
      await adminService.tracerAction({
        TypeObjet: TypeObjetHistorique.Activite,
        ObjetId: activiteId,
        Action: action,
        AncienneValeur: ancienneValeur ? JSON.stringify(ancienneValeur) : undefined,
        NouvelleValeur: nouvelleValeur ? JSON.stringify(nouvelleValeur) : undefined,
        Commentaire: commentaire
      });
    } catch {
      // L'historisation ne doit jamais bloquer l'action métier principale
    }
  };

  const ouvrirNouveau = (parentId?: number): void => {
    setEdition({
      ProjetId: projetId,
      ParentActiviteId: parentId,
      TypeActivite: TypeActivite.Tache,
      StatutActivite: StatutActivite.NonDemarree,
      EstJalon: false,
      AutoriseSaisieTemps: true,
      EstActive: true,
      ModifiableParRessource: false,
      Ordre: activites.length + 1
    });
    setPanelOuvert(true);
  };

  const ouvrirEdition = (activite: IActivity): void => {
    setEdition({ ...activite });
    setPanelOuvert(true);
  };

  const enregistrer = async (): Promise<void> => {
    try {
      if (edition.Id) {
        const avant = activites.filter(a => a.Id === edition.Id)[0];
        await activityService.update(edition.Id, edition);
        await tracerHistorique(ActionHistorique.Modification, edition.Id, avant as unknown as Record<string, unknown>, edition as Record<string, unknown>);
      } else {
        const nouvelle = await activityService.create(edition);
        await tracerHistorique(ActionHistorique.Creation, nouvelle.Id as number, undefined, nouvelle as unknown as Record<string, unknown>);
      }
      setPanelOuvert(false);
      if (projetId) { await chargerActivites(projetId); }
    } catch (e) {
      setErreur("Erreur lors de l'enregistrement : " + (e as Error).message);
    }
  };

  const supprimer = async (activite: IActivity): Promise<void> => {
    if (!window.confirm(`Supprimer définitivement l'activité "${activite.Title}" ?`)) { return; }
    try {
      await activityService.delete(activite.Id as number);
      await tracerHistorique(ActionHistorique.Modification, activite.Id as number, activite as unknown as Record<string, unknown>, undefined, "Suppression de l'activité");
      if (projetId) { await chargerActivites(projetId); }
    } catch (e) {
      setErreur("Erreur lors de la suppression : " + (e as Error).message);
    }
  };

  const modifierChampEnLigne = async (activite: IActivity, champ: keyof IActivity, valeur: string): Promise<void> => {
    const ancienneValeur = { [champ]: activite[champ] };
    const nouvelleValeur = { [champ]: valeur };
    try {
      await activityService.update(activite.Id as number, { [champ]: valeur } as Partial<IActivity>);
      await tracerHistorique(ActionHistorique.Modification, activite.Id as number, ancienneValeur, nouvelleValeur, `Modification directe du champ ${champ}`);
      if (projetId) { await chargerActivites(projetId); }
    } catch (e) {
      setErreur("Erreur lors de la mise à jour : " + (e as Error).message);
    }
  };

  const onDragStart = (item: IActivityNode): void => setLigneGlissee(item);

  const onDrop = async (cible: IActivityNode): Promise<void> => {
    if (!ligneGlissee || ligneGlissee.Id === cible.Id) { setLigneGlissee(undefined); return; }
    if (ligneGlissee.ParentActiviteId !== cible.ParentActiviteId) {
      setErreur("Le glisser-déposer n'est possible qu'entre activités de même niveau (même parent).");
      setLigneGlissee(undefined);
      return;
    }
    try {
      const freresSoeurs = activites
        .filter(a => a.ParentActiviteId === cible.ParentActiviteId)
        .sort((a, b) => (a.Ordre || 0) - (b.Ordre || 0));

      const indexSource = freresSoeurs.findIndex(a => a.Id === ligneGlissee.Id);
      const indexCible = freresSoeurs.findIndex(a => a.Id === cible.Id);
      if (indexSource === -1 || indexCible === -1) { return; }

      const reordonne = [...freresSoeurs];
      const [element] = reordonne.splice(indexSource, 1);
      reordonne.splice(indexCible, 0, element);

      const misesAJour = reordonne.map((a, index) => ({ id: a.Id as number, ordre: index + 1 }));
      for (const mise of misesAJour) {
        await activityService.update(mise.id, { Ordre: mise.ordre });
      }
      await tracerHistorique(
        ActionHistorique.Modification,
        ligneGlissee.Id as number,
        { Ordre: ligneGlissee.Ordre },
        { Ordre: misesAJour.filter(m => m.id === ligneGlissee.Id)[0]?.ordre },
        "Réordonnancement par glisser-déposer"
      );

      if (projetId) { await chargerActivites(projetId); }
    } catch (e) {
      setErreur("Erreur lors du réordonnancement : " + (e as Error).message);
    } finally {
      setLigneGlissee(undefined);
    }
  };

  const onRenderRow: IRenderFunction<IDetailsRowProps> = (rowProps, defaultRender) => {
    if (!rowProps || !defaultRender) { return null; }
    const item: IActivityNode = rowProps.item;
    return (
      <div
        draggable
        onDragStart={() => onDragStart(item)}
        onDragOver={(e) => e.preventDefault()}
        onDrop={() => onDrop(item)}
        style={{ cursor: "grab", opacity: ligneGlissee?.Id === item.Id ? 0.4 : 1 }}
        title="Glisser pour réordonner"
      >
        {defaultRender(rowProps)}
      </div>
    );
  };

  const columns: IColumn[] = [
    {
      key: "poignee", name: "", minWidth: 24, maxWidth: 24,
      onRender: () => <Icon iconName="GripperDotsVertical" styles={{ root: { color: "#605e5c" } }} />
    },
    {
      key: "wbsTitre", name: "WBS / Activité", minWidth: 280, maxWidth: 420,
      onRender: (item: IActivityNode) => (
        <span style={{ paddingLeft: item.depth * INDENT_PAR_NIVEAU, display: "inline-block" }}>
          {item.depth > 0 && <span style={{ color: "#a19f9d", marginRight: 6 }}>└</span>}
          <strong>{item.WBS || ""}</strong> {item.Title}
        </span>
      )
    },
    { key: "type", name: "Type", fieldName: "TypeActivite", minWidth: 70, maxWidth: 90 },
    { key: "statut", name: "Statut", fieldName: "StatutActivite", minWidth: 100, maxWidth: 130 },
    {
      key: "avancement", name: "Avancement", minWidth: 140, maxWidth: 180,
      onRender: (item: IActivityNode) => (
        <ProgressIndicator
          percentComplete={(item.PourcentageAchevement || 0) / 100}
          label={`${item.PourcentageAchevement ?? 0} %`}
          styles={{ progressBar: { background: BIAT_TEAL } }}
        />
      )
    },
    { key: "debutPrevu", name: "Début prévu", minWidth: 90, maxWidth: 110, onRender: (i: IActivityNode) => i.DateDebutPrevue ? new Date(i.DateDebutPrevue).toLocaleDateString() : "-" },
    { key: "finPrevue", name: "Fin prévue", minWidth: 90, maxWidth: 110, onRender: (i: IActivityNode) => i.DateFinPrevue ? new Date(i.DateFinPrevue).toLocaleDateString() : "-" },
    {
      key: "debutReel", name: "Début réel", minWidth: 130, maxWidth: 150,
      onRender: (item: IActivityNode) => (
        <TextField
          type="date"
          value={item.DateDebutReelle ? item.DateDebutReelle.substring(0, 10) : ""}
          onChange={(_, v) => modifierChampEnLigne(item, "DateDebutReelle", v || "").catch(() => { /* géré */ })}
        />
      )
    },
    {
      key: "finReelle", name: "Fin réelle", minWidth: 130, maxWidth: 150,
      onRender: (item: IActivityNode) => (
        <TextField
          type="date"
          value={item.DateFinReelle ? item.DateFinReelle.substring(0, 10) : ""}
          onChange={(_, v) => modifierChampEnLigne(item, "DateFinReelle", v || "").catch(() => { /* géré */ })}
        />
      )
    },
    { key: "charge", name: "Charge (h)", fieldName: "ChargePlanifieeHeures", minWidth: 80, maxWidth: 100 },
    {
      key: "actions", name: "Actions", minWidth: 220, maxWidth: 260,
      onRender: (item: IActivityNode) => (
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <DefaultButton text="Modifier" styles={biatDefaultBtn} onClick={() => ouvrirEdition(item)} />
          <DefaultButton text="Supprimer" styles={biatDefaultBtn} onClick={() => supprimer(item)} />
          {item.TypeActivite !== TypeActivite.Tache &&
            <DefaultButton text="+ Enfant" styles={biatDefaultBtn} onClick={() => ouvrirNouveau(item.Id)} />}
        </Stack>
      )
    }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title="Planning projet" subtitle="Structure WBS, jalons et suivi de l'avancement" />
      <div style={biatCardStyle}>
        {erreur && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErreur("")} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>{erreur}</MessageBar>}

        <div style={biatToolbarStyle}>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
            <Dropdown
              label="Projet"
              options={projets.map(p => ({ key: p.Id as number, text: `${p.CodeProjet} - ${p.Title}` }))}
              selectedKey={projetId}
              onChange={(_, opt) => setProjetId(opt?.key as number)}
              styles={{ root: { width: 320 } }}
            />
            <PrimaryButton text="Nouvelle activité" styles={biatPrimaryBtn} onClick={() => ouvrirNouveau()} disabled={!projetId} />
            {projetId && (
              <DefaultButton
                text={ganttVisible ? "Masquer le Gantt" : "Afficher le Gantt"}
                iconProps={{ iconName: ganttVisible ? "Hide" : "BarChartVertical" }}
                styles={biatDefaultBtn}
                onClick={() => setGanttVisible(v => !v)}
              />
            )}
          </Stack>
        </div>

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}>
            <Spinner label="Chargement du planning..." />
          </div>
        ) : (
          projetId && (
            <>
              {/* Layout horizontal : planning + Gantt côte à côte */}
              <div style={{ display: "flex", flexDirection: "row", alignItems: "flex-start", minWidth: 0 }}>
                {/* Tableau planning — flex 1 */}
                <div style={{ flex: 1, minWidth: 0, overflow: "auto" }}>
                  <DetailsList
                    items={listeAplatie}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    setKey="Id"
                    onRenderRow={onRenderRow}
                  />
                </div>

                {/* Panneau Gantt repliable */}
                {ganttVisible && (
                  <>
                    {/* Divider redimensionnable */}
                    <div
                      onMouseDown={(e) => {
                        isDragging.current = true;
                        dragStart.current = { x: e.clientX, width: ganttWidth };
                        e.preventDefault();
                      }}
                      style={{
                        width: 6,
                        flexShrink: 0,
                        alignSelf: "stretch",
                        cursor: "col-resize",
                        background: "#dce5f0",
                        borderLeft: "1px solid #c8d6e8",
                        borderRight: "1px solid #c8d6e8",
                        transition: "background 0.15s",
                      }}
                      onMouseEnter={e => { (e.currentTarget as HTMLDivElement).style.background = "#b0c4de"; }}
                      onMouseLeave={e => { (e.currentTarget as HTMLDivElement).style.background = "#dce5f0"; }}
                    />

                    {/* Panneau Gantt */}
                    <div style={{ width: ganttWidth, flexShrink: 0, overflow: "hidden", display: "flex", flexDirection: "column" }}>
                      {/* En-tête du panneau */}
                      <div style={{
                        display: "flex", alignItems: "center", justifyContent: "space-between",
                        padding: "6px 10px", background: "#F2F5F9",
                        borderBottom: `2px solid ${BIAT_NAVY}`, flexShrink: 0
                      }}>
                        <span style={{ fontWeight: 700, fontSize: 13, color: BIAT_NAVY }}>Diagramme de Gantt</span>
                        <IconButton
                          iconProps={{ iconName: "ChromeClose" }}
                          title="Fermer"
                          styles={{ root: { width: 28, height: 28, color: BIAT_NAVY } }}
                          onClick={() => setGanttVisible(false)}
                        />
                      </div>

                      {/* Corps scrollable */}
                      <div style={{ overflowX: "auto", overflowY: "auto", flex: 1 }}>
                        <GanttChart
                          lignes={listeAplatie}
                          indentParNiveau={INDENT_PAR_NIVEAU}
                          showLabels={true}
                          timelineWidth={ganttWidth - 20}
                        />
                      </div>
                    </div>
                  </>
                )}
              </div>

              <ComparaisonPlanning lignes={listeAplatie} activityService={activityService} indentParNiveau={INDENT_PAR_NIVEAU} />
            </>
          )
        )}
      </div>

      <Panel
        isOpen={panelOuvert}
        onDismiss={() => setPanelOuvert(false)}
        type={PanelType.medium}
        headerText={edition.Id ? "Modifier l'activité" : "Nouvelle activité"}
        styles={{
          headerText: { color: BIAT_NAVY, fontWeight: '700', fontSize: 16 },
          header: { paddingTop: 20, paddingBottom: 12, borderBottom: `2px solid ${BIAT_TEAL}` }
        }}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Enregistrer" styles={biatPrimaryBtn} onClick={enregistrer} />
            <DefaultButton text="Annuler" styles={biatDefaultBtn} onClick={() => setPanelOuvert(false)} />
          </Stack>
        )}
        isFooterAtBottom
      >
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { paddingTop: 16 } }}>
          <TextField label="Nom" value={edition.Title || ""} onChange={(_, v) => setEdition({ ...edition, Title: v })} />
          <TextField label="Code activité" value={edition.CodeActivite || ""} onChange={(_, v) => setEdition({ ...edition, CodeActivite: v })} />
          <Dropdown
            label="Type"
            options={TYPE_OPTIONS}
            selectedKey={edition.TypeActivite}
            onChange={(_, opt) => setEdition({ ...edition, TypeActivite: opt?.key as TypeActivite })}
          />
          <Dropdown
            label="Statut"
            options={STATUT_OPTIONS}
            selectedKey={edition.StatutActivite}
            onChange={(_, opt) => setEdition({ ...edition, StatutActivite: opt?.key as StatutActivite })}
          />
          <TextField
            label="Charge planifiée (heures)" type="number"
            value={String(edition.ChargePlanifieeHeures ?? "")}
            onChange={(_, v) => setEdition({ ...edition, ChargePlanifieeHeures: Number(v) })}
          />
          <TextField
            label="% Avancement" type="number"
            value={String(edition.PourcentageAchevement ?? "")}
            onChange={(_, v) => setEdition({ ...edition, PourcentageAchevement: Number(v) })}
          />
          <TextField
            label="Date début prévue" type="date"
            value={edition.DateDebutPrevue ? edition.DateDebutPrevue.substring(0, 10) : ""}
            onChange={(_, v) => setEdition({ ...edition, DateDebutPrevue: v })}
          />
          <TextField
            label="Date fin prévue" type="date"
            value={edition.DateFinPrevue ? edition.DateFinPrevue.substring(0, 10) : ""}
            onChange={(_, v) => setEdition({ ...edition, DateFinPrevue: v })}
          />
          <TextField label="Description" multiline value={edition.Description || ""} onChange={(_, v) => setEdition({ ...edition, Description: v })} />
        </Stack>
      </Panel>
    </div>
  );
};

export default PlanningProjet;
