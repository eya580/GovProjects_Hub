import * as React from "react";
import { useEffect, useState, useMemo } from "react";
import {
  DetailsList, IColumn, SelectionMode, Dropdown, IDropdownOption,
  TextField, Stack, Spinner, MessageBar, MessageBarType, Icon,
  PrimaryButton, DefaultButton, Panel, PanelType
} from "@fluentui/react";
import { IPortefeuilleProjetsProps } from "./IPortefeuilleProjetsProps";
import { ProjectService } from "../../../../services/ProjectService";
import { IProject } from "../../../../models/IProject";
import { StatutProjet, Priorite } from "../../../../models/Enums";
import BiatHeader, {
  biatContainerStyle, biatCardStyle, biatToolbarStyle,
  biatPrimaryBtn, biatDefaultBtn, BIAT_NAVY, BIAT_TEAL
} from "../../components/BiatHeader";

const STATUT_OPTIONS: IDropdownOption[] = [
  { key: "", text: "Tous les statuts" },
  ...Object.values(StatutProjet).map(s => ({ key: s, text: s }))
];

const STATUT_OPTIONS_FORM: IDropdownOption[] = Object.values(StatutProjet).map(s => ({ key: s, text: s }));
const PRIORITE_OPTIONS_FORM: IDropdownOption[] = Object.values(Priorite).map(p => ({ key: p, text: p }));

const PROJET_VIDE: Partial<IProject> = {
  StatutProjet: StatutProjet.Brouillon,
  Priorite: Priorite.Normale,
  Actif: true,
  PourcentageAchevement: 0
};

const PortefeuilleProjets: React.FC<IPortefeuilleProjetsProps> = (props) => {
  const service = useMemo(() => new ProjectService(props.context), [props.context]);

  const [projets, setProjets] = useState<IProject[]>([]);
  const [chargement, setChargement] = useState<boolean>(true);
  const [erreur, setErreur] = useState<string>("");
  const [filtreStatut, setFiltreStatut] = useState<string>("");
  const [filtreChefProjetId, setFiltreChefProjetId] = useState<number | undefined>(undefined);
  const [filtreTexte, setFiltreTexte] = useState<string>("");
  const [chefs, setChefs] = useState<{ id: number; title: string }[]>([]);

  const [panelOuvert, setPanelOuvert] = useState<boolean>(false);
  const [edition, setEdition] = useState<Partial<IProject>>(PROJET_VIDE);
  const [enregistrement, setEnregistrement] = useState<boolean>(false);

  const chargerProjets = async (): Promise<void> => {
    setChargement(true);
    setErreur("");
    try {
      const data = await service.getAll(filtreStatut || undefined, filtreChefProjetId);
      setProjets(data);
    } catch (e) {
      setErreur("Erreur lors du chargement des projets : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    chargerProjets().catch(() => { /* géré par setErreur */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filtreStatut, filtreChefProjetId]);

  useEffect(() => {
    const chargerChefs = async (): Promise<void> => {
      try {
        const data = await service.getAll();
        const uniqueChefs = new Map<number, string>();
        data.forEach(p => {
          if (p.ChefProjetId && p.ChefProjet?.Title) {
            uniqueChefs.set(p.ChefProjetId, p.ChefProjet.Title);
          }
        });
        setChefs(Array.from(uniqueChefs.entries()).map(([id, title]) => ({ id, title })));
      } catch (error) {
        console.warn("Impossible de charger les chefs de projet pour les filtres", error);
      }
    };
    chargerChefs().catch(() => undefined);
  }, [service]);

  const estEnRetard = (p: IProject): boolean => {
    if (!p.DateFinPrevue) { return false; }
    const fin = new Date(p.DateFinPrevue);
    return fin < new Date() && p.StatutProjet !== StatutProjet.Cloture;
  };

  const ouvrirNouveauProjet = (): void => {
    setEdition(PROJET_VIDE);
    setPanelOuvert(true);
  };

  const enregistrerProjet = async (): Promise<void> => {
    if (!edition.Title || !edition.CodeProjet) {
      setErreur("Le nom et le code du projet sont obligatoires.");
      return;
    }
    setEnregistrement(true);
    setErreur("");
    try {
      await service.create(edition);
      setPanelOuvert(false);
      await chargerProjets();
    } catch (e) {
      setErreur("Erreur lors de la création du projet : " + (e as Error).message);
    } finally {
      setEnregistrement(false);
    }
  };

  const projetsFiltres = projets.filter(p =>
    !filtreTexte || p.Title.toLowerCase().indexOf(filtreTexte.toLowerCase()) > -1
    || (p.CodeProjet || "").toLowerCase().indexOf(filtreTexte.toLowerCase()) > -1
  );

  const columns: IColumn[] = [
    {
      key: "alerte", name: "", minWidth: 24, maxWidth: 24,
      onRender: (item: IProject) => estEnRetard(item)
        ? <Icon iconName="Warning" title="Projet en retard" style={{ color: "#a4262c" }} /> : null
    },
    { key: "code", name: "Code", fieldName: "CodeProjet", minWidth: 80, maxWidth: 100 },
    {
      key: "titre", name: "Projet", fieldName: "Title", minWidth: 180, maxWidth: 260,
      onRender: (item: IProject) => (
        <a
          href="#"
          onClick={(e) => { e.preventDefault(); props.onOpenProjet?.(item); }}
          style={{ color: BIAT_NAVY, fontWeight: '600', textDecoration: 'none' }}
          onMouseEnter={(e) => { (e.currentTarget as HTMLAnchorElement).style.color = BIAT_TEAL; }}
          onMouseLeave={(e) => { (e.currentTarget as HTMLAnchorElement).style.color = BIAT_NAVY; }}
        >
          {item.Title}
        </a>
      )
    },
    { key: "chef", name: "Chef de projet", minWidth: 140, maxWidth: 180, onRender: (i: IProject) => i.ChefProjet?.Title || "-" },
    { key: "direction", name: "Direction", fieldName: "Direction", minWidth: 100, maxWidth: 140 },
    { key: "statut", name: "Statut", fieldName: "StatutProjet", minWidth: 100, maxWidth: 130 },
    { key: "priorite", name: "Priorité", fieldName: "Priorite", minWidth: 80, maxWidth: 100 },
    {
      key: "avancement", name: "Avancement", minWidth: 100, maxWidth: 120,
      onRender: (item: IProject) => `${item.PourcentageAchevement ?? 0} %`
    },
    { key: "finPrevue", name: "Fin prévue", minWidth: 100, maxWidth: 120, onRender: (i: IProject) => i.DateFinPrevue ? new Date(i.DateFinPrevue).toLocaleDateString() : "-" }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title={props.titre || "Portefeuille de projets"} subtitle="Vue consolidée et suivi de tous les projets" />
      <div style={biatCardStyle}>
        {erreur && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>{erreur}</MessageBar>}

        <div style={biatToolbarStyle}>
          <Stack horizontal tokens={{ childrenGap: 12 }} wrap verticalAlign="center">
            <TextField
              placeholder="Rechercher un projet..."
              value={filtreTexte}
              onChange={(_, v) => setFiltreTexte(v || "")}
              styles={{ root: { width: 260 } }}
            />
            <Dropdown
              options={STATUT_OPTIONS}
              selectedKey={filtreStatut}
              onChange={(_, opt) => setFiltreStatut(opt?.key as string)}
              styles={{ root: { width: 200 } }}
            />
            <Dropdown
              options={[
                { key: "", text: "Tous les chefs" },
                ...chefs.map(c => ({ key: c.id, text: c.title }))
              ]}
              selectedKey={filtreChefProjetId ?? ""}
              onChange={(_, opt) => setFiltreChefProjetId(opt?.key === "" ? undefined : Number(opt?.key))}
              styles={{ root: { width: 240 } }}
            />
            <PrimaryButton text="Nouveau projet" styles={biatPrimaryBtn} iconProps={{ iconName: "Add" }} onClick={ouvrirNouveauProjet} />
          </Stack>
        </div>

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}>
            <Spinner label="Chargement des projets..." />
          </div>
        ) : (
          <DetailsList
            items={projetsFiltres}
            columns={columns}
            selectionMode={SelectionMode.none}
            setKey="Id"
          />
        )}
      </div>

      <Panel
        isOpen={panelOuvert}
        onDismiss={() => setPanelOuvert(false)}
        type={PanelType.medium}
        headerText="Nouveau projet"
        styles={{
          headerText: { color: BIAT_NAVY, fontWeight: '700', fontSize: 16 },
          header: { paddingTop: 20, paddingBottom: 12, borderBottom: `2px solid ${BIAT_TEAL}` }
        }}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Enregistrer" styles={biatPrimaryBtn} onClick={enregistrerProjet} disabled={enregistrement} />
            <DefaultButton text="Annuler" styles={biatDefaultBtn} onClick={() => setPanelOuvert(false)} disabled={enregistrement} />
          </Stack>
        )}
        isFooterAtBottom
      >
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { paddingTop: 16 } }}>
          <TextField
            label="Nom du projet" required
            value={edition.Title || ""}
            onChange={(_, v) => setEdition({ ...edition, Title: v })}
          />
          <TextField
            label="Code projet" required
            value={edition.CodeProjet || ""}
            onChange={(_, v) => setEdition({ ...edition, CodeProjet: v })}
          />
          <TextField
            label="Description" multiline rows={3}
            value={edition.Description || ""}
            onChange={(_, v) => setEdition({ ...edition, Description: v })}
          />
          <TextField
            label="Direction"
            value={edition.Direction || ""}
            onChange={(_, v) => setEdition({ ...edition, Direction: v })}
          />
          <Dropdown
            label="Statut"
            options={STATUT_OPTIONS_FORM}
            selectedKey={edition.StatutProjet}
            onChange={(_, opt) => setEdition({ ...edition, StatutProjet: opt?.key as StatutProjet })}
          />
          <Dropdown
            label="Priorité"
            options={PRIORITE_OPTIONS_FORM}
            selectedKey={edition.Priorite}
            onChange={(_, opt) => setEdition({ ...edition, Priorite: opt?.key as Priorite })}
          />
          <TextField
            label="Avancement (%)" type="number"
            value={edition.PourcentageAchevement !== undefined ? String(edition.PourcentageAchevement) : ""}
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
          <TextField
            label="Budget" type="number"
            value={String(edition.Budget ?? "")}
            onChange={(_, v) => setEdition({ ...edition, Budget: Number(v) })}
          />
        </Stack>
      </Panel>
    </div>
  );
};

export default PortefeuilleProjets;
