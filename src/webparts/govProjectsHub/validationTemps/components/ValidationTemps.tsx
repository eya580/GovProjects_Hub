import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  DetailsList, IColumn, SelectionMode, Stack, Spinner, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, Panel, PanelType, TextField
} from "@fluentui/react";
import { IValidationTempsProps } from "./IValidationTempsProps";
import { TimesheetService } from "../../../../services/TimesheetService";
import { ITimesheet } from "../../../../models/ITimesheet";
import BiatHeader, {
  biatContainerStyle, biatCardStyle,
  biatPrimaryBtn, biatDefaultBtn, biatDangerBtn, BIAT_NAVY, BIAT_TEAL
} from "../../components/BiatHeader";

const ValidationTemps: React.FC<IValidationTempsProps> = (props) => {
  const timesheetService = useMemo(() => new TimesheetService(props.context), [props.context]);

  const [feuilles, setFeuilles] = useState<ITimesheet[]>([]);
  const [chargement, setChargement] = useState<boolean>(true);
  const [erreur, setErreur] = useState<string>("");
  const [panelOuvert, setPanelOuvert] = useState<boolean>(false);
  const [feuilleSelectionnee, setFeuilleSelectionnee] = useState<ITimesheet>();
  const [commentaire, setCommentaire] = useState<string>("");

  const charger = async (): Promise<void> => {
    setChargement(true);
    setErreur("");
    try {
      const data = await timesheetService.getFeuillesEnAttente();
      setFeuilles(data);
    } catch (e) {
      setErreur("Erreur lors du chargement des feuilles en attente : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    charger().catch(() => { /* géré */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const valider = async (feuille: ITimesheet): Promise<void> => {
    try {
      const validateurId = (props.context.pageContext.legacyPageContext as { userId?: number })?.userId || 0;
      await timesheetService.valider(feuille.Id as number, validateurId);
      await charger();
    } catch (e) {
      setErreur("Erreur lors de la validation : " + (e as Error).message);
    }
  };

  const ouvrirRejet = (feuille: ITimesheet): void => {
    setFeuilleSelectionnee(feuille);
    setCommentaire("");
    setPanelOuvert(true);
  };

  const confirmerRejet = async (): Promise<void> => {
    if (!feuilleSelectionnee) { return; }
    try {
      await timesheetService.rejeter(feuilleSelectionnee.Id as number, commentaire);
      setPanelOuvert(false);
      await charger();
    } catch (e) {
      setErreur("Erreur lors du rejet : " + (e as Error).message);
    }
  };

  const columns: IColumn[] = [
    { key: "code", name: "Feuille", fieldName: "CodeFeuilleTemps", minWidth: 120, maxWidth: 160 },
    { key: "ressource", name: "Ressource (Id)", fieldName: "RessourceId", minWidth: 90, maxWidth: 110 },
    { key: "total", name: "Total (h)", fieldName: "TotalHeures", minWidth: 70, maxWidth: 90 },
    { key: "projet", name: "Dont projet (h)", fieldName: "TotalHeuresProjet", minWidth: 90, maxWidth: 110 },
    { key: "admin", name: "Dont admin. (h)", fieldName: "TotalHeuresAdministratives", minWidth: 90, maxWidth: 110 },
    { key: "soumission", name: "Soumise le", minWidth: 100, maxWidth: 120, onRender: (i: ITimesheet) => i.DateSoumission ? new Date(i.DateSoumission).toLocaleString() : "-" },
    {
      key: "actions", name: "Actions", minWidth: 200, maxWidth: 240,
      onRender: (item: ITimesheet) => (
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <PrimaryButton text="Valider" styles={biatPrimaryBtn} onClick={() => valider(item)} />
          <DefaultButton text="Rejeter" styles={biatDangerBtn} onClick={() => ouvrirRejet(item)} />
        </Stack>
      )
    }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title="Validation des feuilles de temps" subtitle="Approbation et rejet des feuilles soumises" />
      <div style={biatCardStyle}>
        {erreur && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>{erreur}</MessageBar>}

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}>
            <Spinner label="Chargement des feuilles en attente..." />
          </div>
        ) : (
          <DetailsList items={feuilles} columns={columns} selectionMode={SelectionMode.none} setKey="Id" />
        )}
      </div>

      <Panel
        isOpen={panelOuvert}
        onDismiss={() => setPanelOuvert(false)}
        type={PanelType.medium}
        headerText="Rejeter la feuille de temps"
        styles={{
          headerText: { color: BIAT_NAVY, fontWeight: '700', fontSize: 16 },
          header: { paddingTop: 20, paddingBottom: 12, borderBottom: `2px solid ${BIAT_TEAL}` }
        }}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Confirmer le rejet" styles={biatDangerBtn} onClick={confirmerRejet} />
            <DefaultButton text="Annuler" styles={biatDefaultBtn} onClick={() => setPanelOuvert(false)} />
          </Stack>
        )}
        isFooterAtBottom
      >
        <Stack styles={{ root: { paddingTop: 16 } }}>
          <TextField
            label="Commentaire de rejet"
            multiline
            rows={4}
            value={commentaire}
            onChange={(_, v) => setCommentaire(v || "")}
          />
        </Stack>
      </Panel>
    </div>
  );
};

export default ValidationTemps;
