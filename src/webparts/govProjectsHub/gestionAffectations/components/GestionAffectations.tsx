import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  DetailsList, IColumn, SelectionMode, Stack, Spinner, MessageBar, MessageBarType,
  Dropdown, IDropdownOption, TextField, PrimaryButton, DefaultButton, Panel, PanelType, Icon,
  IconButton
} from "@fluentui/react";
import { IGestionAffectationsProps } from "./IGestionAffectationsProps";
import { ProjectService } from "../../../../services/ProjectService";
import { ActivityService } from "../../../../services/ActivityService";
import { ResourceService } from "../../../../services/ResourceService";
import { AssignmentService } from "../../../../services/AssignmentService";
import { IProject } from "../../../../models/IProject";
import { IActivity } from "../../../../models/IActivity";
import { IResource } from "../../../../models/IResource";
import { IAssignment, IAssignmentView } from "../../../../models/IAssignment";
import { StatutAffectation } from "../../../../models/Enums";
import BiatHeader, {
  biatContainerStyle, biatCardStyle, biatToolbarStyle,
  biatPrimaryBtn, biatDefaultBtn, BIAT_NAVY, BIAT_TEAL
} from "../../components/BiatHeader";

const STATUT_AFFECTATION_OPTIONS: IDropdownOption[] = Object.values(StatutAffectation).map(s => ({ key: s, text: s }));

const GestionAffectations: React.FC<IGestionAffectationsProps> = (props) => {
  const projectService = useMemo(() => new ProjectService(props.context), [props.context]);
  const activityService = useMemo(() => new ActivityService(props.context), [props.context]);
  const resourceService = useMemo(() => new ResourceService(props.context), [props.context]);
  const assignmentService = useMemo(() => new AssignmentService(props.context), [props.context]);

  const [projets, setProjets] = useState<IProject[]>([]);
  const [projetId, setProjetId] = useState<number | undefined>(undefined);
  const [activites, setActivites] = useState<IActivity[]>([]);
  const [ressources, setRessources] = useState<IResource[]>([]);
  const [affectations, setAffectations] = useState<IAssignmentView[]>([]);
  const [chargement, setChargement] = useState<boolean>(false);
  const [erreur, setErreur] = useState<string>("");
  const [panelOuvert, setPanelOuvert] = useState<boolean>(false);
  const [edition, setEdition] = useState<Partial<IAssignment>>({});

  useEffect(() => {
    projectService.getAll().then(setProjets).catch(() => setErreur("Erreur de chargement des projets."));
    resourceService.getAll().then(setRessources).catch(() => setErreur("Erreur de chargement des ressources."));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const charger = async (pId: number): Promise<void> => {
    setChargement(true);
    setErreur("");
    try {
      const [acts, affs] = await Promise.all([
        activityService.getByProjet(pId),
        assignmentService.getByProjet(pId)
      ]);
      setActivites(acts.filter(a => activityService.estSaisissable(a) || true));
      const vues: IAssignmentView[] = affs.map(a => {
        const ressource = ressources.find(r => r.Id === a.RessourceId);
        const activite = acts.find(ac => ac.Id === a.ActiviteId);
        const capaciteRestante = (ressource?.CapaciteHebdomadaireHeures || 0) - (a.ChargeAffecteeHeures || 0);
        return {
          ...a,
          RessourceNom: ressource?.Title,
          ActiviteNom: activite?.Title,
          CapaciteRestanteHeures: capaciteRestante,
          EstEnSurcharge: assignmentService.detecterSurcharge(a.ChargeAffecteeHeures || 0, ressource?.CapaciteHebdomadaireHeures || 0)
        };
      });
      setAffectations(vues);
    } catch (e) {
      setErreur("Erreur lors du chargement des affectations : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    if (projetId) { charger(projetId).catch(() => { /* géré */ }); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projetId, ressources]);

  const ouvrirNouveau = (): void => {
    setEdition({
      ProjetId: projetId,
      StatutAffectation: StatutAffectation.Prevue,
      ResponsableTache: false,
      AutoriseSaisieTemps: true,
      Actif: true
    });
    setPanelOuvert(true);
  };

  const modifierAffectation = (affectation: IAssignment): void => {
    setEdition(affectation);
    setPanelOuvert(true);
  };

  const supprimerAffectation = async (affectation: IAssignment): Promise<void> => {
    if (!affectation.Id) { return; }
    if (!window.confirm(`Supprimer définitivement l'affectation de ${affectation.RessourceId ? affectation.RessourceId : "cette ressource"} ?`)) { return; }
    try {
      await assignmentService.delete(affectation.Id);
      if (projetId) { await charger(projetId); }
    } catch (e) {
      setErreur("Erreur lors de la suppression : " + (e as Error).message);
    }
  };

  const changerStatut = async (affectation: IAssignment, statut: StatutAffectation): Promise<void> => {
    if (!affectation.Id) { return; }
    try {
      await assignmentService.update(affectation.Id, { StatutAffectation: statut });
      if (projetId) { await charger(projetId); }
    } catch (e) {
      setErreur("Erreur lors de la mise à jour du statut : " + (e as Error).message);
    }
  };

  const enregistrer = async (): Promise<void> => {
    try {
      if (edition.Id) {
        await assignmentService.update(edition.Id, edition);
      } else {
        await assignmentService.create(edition);
      }
      setPanelOuvert(false);
      if (projetId) { await charger(projetId); }
    } catch (e) {
      setErreur("Erreur lors de l'enregistrement : " + (e as Error).message);
    }
  };

  const columns: IColumn[] = [
    {
      key: "alerte", name: "", minWidth: 24, maxWidth: 24,
      onRender: (item: IAssignmentView) => item.EstEnSurcharge
        ? <Icon iconName="Warning" title="Ressource en surcharge" style={{ color: "#a4262c" }} /> : null
    },
    { key: "ressource", name: "Ressource", fieldName: "RessourceNom", minWidth: 140, maxWidth: 180 },
    { key: "activite", name: "Tâche", fieldName: "ActiviteNom", minWidth: 160, maxWidth: 220 },
    { key: "role", name: "Rôle", fieldName: "RoleProjet", minWidth: 90, maxWidth: 110 },
    { key: "charge", name: "Charge (h)", fieldName: "ChargeAffecteeHeures", minWidth: 80, maxWidth: 100 },
    { key: "unite", name: "Unité (%)", fieldName: "UniteAffectation", minWidth: 80, maxWidth: 100 },
    { key: "restante", name: "Capacité restante (h)", fieldName: "CapaciteRestanteHeures", minWidth: 130, maxWidth: 160 },
    {
      key: "statut", name: "Statut", minWidth: 150, maxWidth: 180,
      onRender: (item: IAssignmentView) => (
        <Dropdown
          options={STATUT_AFFECTATION_OPTIONS}
          selectedKey={item.StatutAffectation}
          onChange={(_, opt) => opt?.key && changerStatut(item, opt.key as StatutAffectation)}
          styles={{
            root: { width: 140 },
            title: { border: "none", boxShadow: "none", background: "transparent" },
            dropdown: { border: "none" }
          }}
        />
      )
    },
    {
      key: "actions", name: "Actions", minWidth: 90, maxWidth: 110,
      onRender: (item: IAssignmentView) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <IconButton iconProps={{ iconName: "Edit" }} title="Modifier" onClick={() => modifierAffectation(item)} />
          <IconButton iconProps={{ iconName: "Delete" }} title="Supprimer" onClick={() => supprimerAffectation(item)} />
        </Stack>
      )
    },
    { key: "debut", name: "Début", minWidth: 90, maxWidth: 100, onRender: (i: IAssignmentView) => i.DateDebutAffectation ? new Date(i.DateDebutAffectation).toLocaleDateString() : "-" },
    { key: "fin", name: "Fin", minWidth: 90, maxWidth: 100, onRender: (i: IAssignmentView) => i.DateFinAffectation ? new Date(i.DateFinAffectation).toLocaleDateString() : "-" }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title={props.titre || "Gestion des affectations"} subtitle="Affectation des ressources aux tâches de projet" />
      <div style={biatCardStyle}>
        {erreur && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>{erreur}</MessageBar>}

        <div style={biatToolbarStyle}>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
            <Dropdown
              label="Projet"
              options={projets.map(p => ({ key: p.Id as number, text: `${p.CodeProjet} - ${p.Title}` }))}
              selectedKey={projetId}
              onChange={(_, opt) => setProjetId(opt?.key as number)}
              styles={{ root: { width: 320 } }}
            />
            <PrimaryButton text="Nouvelle affectation" styles={biatPrimaryBtn} onClick={ouvrirNouveau} disabled={!projetId} />
          </Stack>
        </div>

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}>
            <Spinner label="Chargement des affectations..." />
          </div>
        ) : (
          projetId && <DetailsList items={affectations} columns={columns} selectionMode={SelectionMode.none} setKey="Id" />
        )}
      </div>

      <Panel
        isOpen={panelOuvert}
        onDismiss={() => setPanelOuvert(false)}
        type={PanelType.medium}
        headerText="Nouvelle affectation"
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
          <Dropdown
            label="Tâche"
            placeholder="Sélectionner une tâche"
            options={activites
              .filter(a => a.Id !== undefined && a.Id !== null)
              .map(a => ({ key: a.Id as number, text: a.Title || "(tâche sans titre)" } as IDropdownOption))}
            selectedKey={edition.ActiviteId ?? undefined}
            onChange={(_, opt) => setEdition({ ...edition, ActiviteId: opt?.key !== undefined ? Number(opt.key) : undefined })}
          />
          <Dropdown
            label="Ressource"
            placeholder="Sélectionner une ressource"
            options={ressources
              .filter(r => r.Id !== undefined && r.Id !== null)
              .map(r => ({ key: r.Id as number, text: r.Title || "(ressource sans nom)" } as IDropdownOption))}
            selectedKey={edition.RessourceId ?? undefined}
            onChange={(_, opt) => setEdition({ ...edition, RessourceId: opt?.key !== undefined ? Number(opt.key) : undefined })}
          />
          <TextField
            label="Charge affectée (heures)" type="number"
            value={String(edition.ChargeAffecteeHeures ?? "")}
            onChange={(_, v) => setEdition({ ...edition, ChargeAffecteeHeures: Number(v) })}
          />
          <TextField
            label="Unité d'affectation (%)" type="number"
            value={String(edition.UniteAffectation ?? "")}
            onChange={(_, v) => setEdition({ ...edition, UniteAffectation: Number(v) })}
          />
          <TextField
            label="Date début" type="date"
            value={edition.DateDebutAffectation ? edition.DateDebutAffectation.substring(0, 10) : ""}
            onChange={(_, v) => setEdition({ ...edition, DateDebutAffectation: v })}
          />
          <TextField
            label="Date fin" type="date"
            value={edition.DateFinAffectation ? edition.DateFinAffectation.substring(0, 10) : ""}
            onChange={(_, v) => setEdition({ ...edition, DateFinAffectation: v })}
          />
        </Stack>
      </Panel>
    </div>
  );
};

export default GestionAffectations;
