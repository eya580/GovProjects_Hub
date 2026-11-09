import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  Pivot, PivotItem, DetailsList, IColumn, SelectionMode, Stack, Spinner,
  MessageBar, MessageBarType, PrimaryButton, DefaultButton, TextField, Toggle, Panel, PanelType, Dropdown
} from "@fluentui/react";
import { IAdministrationProps } from "./IAdministrationProps";
import { AdminService } from "../../../../services/AdminService";
import { IAdminTaskType, ITimesheetPeriod } from "../../../../models/ITimesheet";
import { IApplicationSetting } from "../../../../models/IAdministration";
import { StatutPeriode } from "../../../../models/Enums";
import BiatHeader, {
  biatContainerStyle, biatCardStyle, biatToolbarStyle,
  biatPrimaryBtn, biatDefaultBtn, BIAT_NAVY, BIAT_TEAL
} from "../../components/BiatHeader";

type IParametreLigne = IApplicationSetting & { estNouveau?: boolean };
const Administration: React.FC<IAdministrationProps> = (props) => {
  const adminService = useMemo(() => new AdminService(props.context), [props.context]);

  const [typesAdmin, setTypesAdmin] = useState<IAdminTaskType[]>([]);
  const [periodes, setPeriodes] = useState<ITimesheetPeriod[]>([]);
  const [parametres, setParametres] = useState<IParametreLigne[]>([]);
  const [chargement, setChargement] = useState<boolean>(true);
  const [erreur, setErreur] = useState<string>("");

  const [panelTypeOuvert, setPanelTypeOuvert] = useState<boolean>(false);
  const [editionType, setEditionType] = useState<Partial<IAdminTaskType>>({});

  const [nbSemaines, setNbSemaines] = useState<string>("52");
  const [anneeDebut, setAnneeDebut] = useState<string>(String(new Date().getFullYear()));
  const [dateDebut, setDateDebut] = useState<string>(new Date().toISOString().substring(0, 10));
  const [semaineDebut, setSemaineDebut] = useState<string>("1");
  const statutPeriode = StatutPeriode.Preparee;

  const chargerTout = async (): Promise<void> => {
    setChargement(true);
    setErreur("");
    try {
      const [t, p, params] = await Promise.all([
        adminService.getTypesTachesAdministratives(false),
        adminService.getPeriodes(),
        adminService.getParametres()
      ]);
      setTypesAdmin(t);
      setPeriodes(p);
      setParametres(params);
    } catch (e) {
      setErreur("Erreur lors du chargement : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    chargerTout().catch(() => { /* géré */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const ouvrirNouveauType = (): void => {
    setEditionType({ NecessiteValidation: false, EstAbsence: false, ImpacteCapacite: false, SaisieHeuresAutorisee: true, SaisieJourneeAutorisee: false, Actif: true, Ordre: typesAdmin.length + 1 });
    setPanelTypeOuvert(true);
  };

  const enregistrerType = async (): Promise<void> => {
    try {
      await adminService.saveTypeTacheAdministrative(editionType);
      const typeSauvegarde = {
        ...(editionType as IAdminTaskType),
        Id: editionType.Id ?? Date.now(),
        Actif: editionType.Actif ?? true
      } as IAdminTaskType;

      setTypesAdmin(prev => {
        if (editionType.Id) {
          return prev.map(type => type.Id === editionType.Id ? { ...type, ...editionType } as IAdminTaskType : type);
        }
        return [typeSauvegarde, ...prev];
      });
      setPanelTypeOuvert(false);
      setEditionType({});
      setErreur("");
    } catch (e) {
      setErreur("Erreur lors de l'enregistrement : " + (e as Error).message);
    }
  };

  const genererPeriodes = async (): Promise<void> => {
    try {
      const nouvelles = adminService.genererPeriodesHebdomadaires(
        Number(anneeDebut),
        Number(semaineDebut),
        Number(nbSemaines),
        dateDebut,
        statutPeriode
      );
      for (const p of nouvelles) { await adminService.savePeriode(p); }
      const periodesAjoutees = nouvelles.map((p, index) => ({ ...(p as ITimesheetPeriod), Id: Date.now() + index } as ITimesheetPeriod));
      setPeriodes(prev => [...prev, ...periodesAjoutees]);
      setErreur("");
    } catch (e) {
      setErreur("Erreur lors de la génération des périodes : " + (e as Error).message);
    }
  };

  const ajouterParametre = (): void => {
    setParametres(prev => [{ Cle: "", Valeur: "", TypeValeur: "", Description: "", Actif: true, estNouveau: true }, ...prev]);
  };

  const mettreAJourParametre = (index: number, champ: keyof IApplicationSetting, valeur: string | boolean): void => {
    setParametres(prev => prev.map((param, idx) => idx === index ? { ...param, [champ]: valeur } : param));
  };

  const enregistrerParametre = async (param: IParametreLigne): Promise<void> => {
    const cle = param.Cle?.trim() || "";
    const valeur = param.Valeur?.trim() || "";
    const type = param.TypeValeur?.trim() || "";

    if (!cle || !valeur || !type) {
      setErreur("Veuillez remplir la clé, la valeur et le type du paramètre.");
      return;
    }

    try {
      await adminService.saveParametre({
        ...param,
        Cle: cle,
        Valeur: valeur,
        TypeValeur: type,
        Description: param.Description?.trim() || "",
        Actif: param.Actif
      });
      setParametres(prev => prev.map(item => item === param ? {
        ...item,
        Cle: cle,
        Valeur: valeur,
        TypeValeur: type,
        Description: param.Description?.trim() || "",
        Actif: param.Actif,
        estNouveau: false
      } : item));
      setErreur("");
    } catch (e) {
      setErreur("Erreur lors de l'enregistrement du paramètre : " + (e as Error).message);
    }
  };

  const modifierStatutPeriode = async (periode: ITimesheetPeriod, statut: StatutPeriode): Promise<void> => {
    try {
      await adminService.savePeriode({ ...periode, StatutPeriode: statut });
      setPeriodes(prev => prev.map(item => item.Id === periode.Id ? { ...item, StatutPeriode: statut } : item));
      setErreur("");
    } catch (e) {
      setErreur("Erreur lors de la mise à jour du statut : " + (e as Error).message);
    }
  };

  const supprimerTypeAdmin = async (item: IAdminTaskType): Promise<void> => {
    try {
      if (item.Id) {
        await adminService.deleteTypeTacheAdministrative(item.Id);
      }
      setTypesAdmin(prev => prev.filter(type => type !== item));
      setErreur("");
    } catch (e) {
      setErreur("Erreur lors de la suppression du type : " + (e as Error).message);
    }
  };

  const supprimerParametre = async (item: IParametreLigne): Promise<void> => {
    try {
      if (item.Id) {
        await adminService.deleteParametre(item.Id);
      }
      setParametres(prev => prev.filter(param => param !== item));
      setErreur("");
    } catch (e) {
      setErreur("Erreur lors de la suppression du paramètre : " + (e as Error).message);
    }
  };

  const colonnesTypes: IColumn[] = [
    { key: "titre", name: "Libellé", fieldName: "Title", minWidth: 140, maxWidth: 180 },
    { key: "code", name: "Code", fieldName: "CodeType", minWidth: 90, maxWidth: 110 },
    { key: "absence", name: "Absence", minWidth: 70, maxWidth: 80, onRender: (i: IAdminTaskType) => i.EstAbsence ? "Oui" : "Non" },
    { key: "impacte", name: "Impacte capacité", minWidth: 100, maxWidth: 120, onRender: (i: IAdminTaskType) => i.ImpacteCapacite ? "Oui" : "Non" },
    { key: "validation", name: "Validation requise", minWidth: 110, maxWidth: 130, onRender: (i: IAdminTaskType) => i.NecessiteValidation ? "Oui" : "Non" },
    { key: "actif", name: "Actif", minWidth: 60, maxWidth: 70, onRender: (i: IAdminTaskType) => i.Actif ? "Oui" : "Non" },
    {
      key: "actions", name: "", minWidth: 180, maxWidth: 190,
      onRender: (item: IAdminTaskType) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Modifier" styles={biatDefaultBtn} onClick={() => { setEditionType(item); setPanelTypeOuvert(true); }} />
          <DefaultButton text="Supprimer" styles={biatDefaultBtn} onClick={() => supprimerTypeAdmin(item).catch(() => { /* géré */ })} />
        </Stack>
      )
    }
  ];

  const colonnesPeriodes: IColumn[] = [
    { key: "titre", name: "Période", fieldName: "Title", minWidth: 140, maxWidth: 180 },
    { key: "code", name: "Code", fieldName: "CodePeriode", minWidth: 90, maxWidth: 110 },
    { key: "debut", name: "Début", minWidth: 90, maxWidth: 100, onRender: (i: ITimesheetPeriod) => i.DateDebut ? new Date(i.DateDebut).toLocaleDateString() : "-" },
    { key: "fin", name: "Fin", minWidth: 90, maxWidth: 100, onRender: (i: ITimesheetPeriod) => i.DateFin ? new Date(i.DateFin).toLocaleDateString() : "-" },
    {
      key: "statut", name: "Statut", minWidth: 120, maxWidth: 150,
      onRender: (item: ITimesheetPeriod) => (
        <Dropdown
          options={Object.values(StatutPeriode).map(s => ({ key: s, text: s }))}
          selectedKey={item.StatutPeriode}
          onChange={(_, opt) => opt?.key && modifierStatutPeriode(item, opt.key as StatutPeriode)}
          styles={{ root: { width: 140 }, title: { border: "none", boxShadow: "none", background: "transparent" }, dropdown: { border: "none" } }}
        />
      )
    }
  ];

  const colonnesParametres: IColumn[] = [
    {
      key: "cle", name: "Clé", minWidth: 160, maxWidth: 220,
      onRender: (item: IParametreLigne, index?: number) => (
        <TextField value={item.Cle || ""} onChange={(_, v) => index !== undefined && mettreAJourParametre(index, "Cle", v || "")} />
      )
    },
    {
      key: "valeur", name: "Valeur", minWidth: 180, maxWidth: 240,
      onRender: (item: IParametreLigne, index?: number) => (
        <TextField value={item.Valeur || ""} onChange={(_, v) => index !== undefined && mettreAJourParametre(index, "Valeur", v || "")} />
      )
    },
    {
      key: "type", name: "Type", minWidth: 100, maxWidth: 140,
      onRender: (item: IParametreLigne, index?: number) => (
        <TextField value={item.TypeValeur || ""} onChange={(_, v) => index !== undefined && mettreAJourParametre(index, "TypeValeur", v || "")} />
      )
    },
    {
      key: "description", name: "Description", minWidth: 220, maxWidth: 320,
      onRender: (item: IParametreLigne, index?: number) => (
        <TextField value={item.Description || ""} onChange={(_, v) => index !== undefined && mettreAJourParametre(index, "Description", v || "")} />
      )
    },
    {
      key: "actions", name: "", minWidth: 200, maxWidth: 220,
      onRender: (item: IParametreLigne) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text={item.estNouveau ? "Ajouter" : "Enregistrer"} styles={biatPrimaryBtn} onClick={() => enregistrerParametre(item).catch(() => { /* géré */ })} />
          <DefaultButton text="Supprimer" styles={biatDefaultBtn} onClick={() => supprimerParametre(item).catch(() => { /* géré */ })} />
        </Stack>
      )
    }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title="Administration" subtitle="Gestion des tâches, périodes et paramètres système" />
      <div style={biatCardStyle}>
        {erreur && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>{erreur}</MessageBar>}

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}>
            <Spinner label="Chargement de l'administration..." />
          </div>
        ) : (
          <Pivot
            styles={{
              link: { color: '#4A5568', fontWeight: '500' as const },
              linkIsSelected: [{ color: BIAT_NAVY, fontWeight: '700' as const }, { selectors: { '::before': { backgroundColor: BIAT_TEAL } } }]
            }}
          >
            <PivotItem headerText="Tâches administratives">
              <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 16 } }}>
                <div style={biatToolbarStyle}>
                  <PrimaryButton text="Nouveau type" styles={{ ...biatPrimaryBtn, root: { ...biatPrimaryBtn.root, width: 160 } }} onClick={ouvrirNouveauType} />
                </div>
                <DetailsList items={typesAdmin} columns={colonnesTypes} selectionMode={SelectionMode.none} setKey="Id" />
              </Stack>
            </PivotItem>

            <PivotItem headerText="Périodes de feuille de temps">
              <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 16 } }}>
                <div style={biatToolbarStyle}>
                  <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="end">
                    <TextField label="Année de début" type="number" value={anneeDebut} onChange={(_, v) => setAnneeDebut(v || anneeDebut)} />
                    <TextField label="Date de début" type="date" value={dateDebut} onChange={(_, v) => setDateDebut(v || dateDebut)} />
                    <TextField label="N° semaine de début" type="number" value={semaineDebut} onChange={(_, v) => setSemaineDebut(v || semaineDebut)} />
                    <TextField label="Nombre de semaines" type="number" value={nbSemaines} onChange={(_, v) => setNbSemaines(v || nbSemaines)} />
                    <PrimaryButton text="Générer les périodes" styles={biatPrimaryBtn} onClick={() => genererPeriodes().catch(() => { /* géré */ })} />
                  </Stack>
                </div>
                <DetailsList items={periodes} columns={colonnesPeriodes} selectionMode={SelectionMode.none} setKey="Id" />
              </Stack>
            </PivotItem>

            <PivotItem headerText="Paramètres">
              <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 16 } }}>
                <div style={biatToolbarStyle}>
                  <PrimaryButton text="Ajouter un paramètre" styles={{ ...biatPrimaryBtn, root: { ...biatPrimaryBtn.root, width: 180 } }} onClick={ajouterParametre} />
                </div>
                <DetailsList items={parametres} columns={colonnesParametres} selectionMode={SelectionMode.none} setKey="Id" />
              </Stack>
            </PivotItem>
          </Pivot>
        )}
      </div>

      <Panel
        isOpen={panelTypeOuvert}
        onDismiss={() => setPanelTypeOuvert(false)}
        type={PanelType.medium}
        headerText="Type de tâche administrative"
        styles={{
          headerText: { color: BIAT_NAVY, fontWeight: '700' as const, fontSize: 16 },
          header: { paddingTop: 20, paddingBottom: 12, borderBottom: `2px solid ${BIAT_TEAL}` }
        }}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Enregistrer" styles={biatPrimaryBtn} onClick={() => enregistrerType().catch(() => { /* géré */ })} />
            <DefaultButton text="Annuler" styles={biatDefaultBtn} onClick={() => setPanelTypeOuvert(false)} />
          </Stack>
        )}
        isFooterAtBottom
      >
        <Stack tokens={{ childrenGap: 12 }} styles={{ root: { paddingTop: 16 } }}>
          <TextField label="Libellé" value={editionType.Title || ""} onChange={(_, v) => setEditionType({ ...editionType, Title: v })} />
          <TextField label="Code" value={editionType.CodeType || ""} onChange={(_, v) => setEditionType({ ...editionType, CodeType: v })} />
          <Toggle label="Nécessite validation" checked={!!editionType.NecessiteValidation} onChange={(_, c) => setEditionType({ ...editionType, NecessiteValidation: c })} />
          <Toggle label="Est une absence" checked={!!editionType.EstAbsence} onChange={(_, c) => setEditionType({ ...editionType, EstAbsence: c })} />
          <Toggle label="Impacte la capacité" checked={!!editionType.ImpacteCapacite} onChange={(_, c) => setEditionType({ ...editionType, ImpacteCapacite: c })} />
          <Toggle label="Saisie en heures autorisée" checked={!!editionType.SaisieHeuresAutorisee} onChange={(_, c) => setEditionType({ ...editionType, SaisieHeuresAutorisee: c })} />
          <Toggle label="Saisie en journée autorisée" checked={!!editionType.SaisieJourneeAutorisee} onChange={(_, c) => setEditionType({ ...editionType, SaisieJourneeAutorisee: c })} />
          <Toggle label="Actif" checked={!!editionType.Actif} onChange={(_, c) => setEditionType({ ...editionType, Actif: c })} />
        </Stack>
      </Panel>
    </div>
  );
};

export default Administration;
