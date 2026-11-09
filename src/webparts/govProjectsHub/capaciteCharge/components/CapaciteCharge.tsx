import * as React from "react";
import { useEffect, useMemo, useState } from "react";
import {
  DetailsList, IColumn, SelectionMode, Stack, Spinner, MessageBar, MessageBarType,
  PrimaryButton, DefaultButton, TextField, Icon
} from "@fluentui/react";
import { ICapaciteChargeProps } from "./ICapaciteChargeProps";
import { ResourceService } from "../../../../services/ResourceService";
import { AssignmentService } from "../../../../services/AssignmentService";
import { TimesheetService } from "../../../../services/TimesheetService";
import { IResource } from "../../../../models/IResource";
import BiatHeader, {
  biatContainerStyle, biatCardStyle, biatToolbarStyle,
  biatPrimaryBtn, biatDefaultBtn
} from "../../components/BiatHeader";
interface ILigneCapaciteCharge {
  ressourceId: number;
  ressourceNom: string;
  capaciteTheorique: number;
  heuresAdmin: number;
  capaciteNette: number;
  chargeAffectee: number;
  tauxDeCharge: number;
  enSurcharge: boolean;
}

const calculerJoursOuvres = (debut: string, fin: string): number => {
  let count = 0;
  const startMs   = new Date(debut + "T00:00:00").getTime();
  const endMs     = new Date(fin   + "T00:00:00").getTime();
  const MS_PER_DAY = 86400000;
  for (let ms = startMs; ms <= endMs; ms += MS_PER_DAY) {
    const jour = new Date(ms).getDay();
    if (jour !== 0 && jour !== 6) { count++; }
  }
  return count;
};

const CapaciteCharge: React.FC<ICapaciteChargeProps> = (props) => {
  const resourceService = useMemo(() => new ResourceService(props.context), [props.context]);
  const assignmentService = useMemo(() => new AssignmentService(props.context), [props.context]);
  const timesheetService = useMemo(() => new TimesheetService(props.context), [props.context]);

  const [dateDebut, setDateDebut] = useState<string>(new Date().toISOString().substring(0, 10));
  const [dateFin, setDateFin] = useState<string>(new Date(Date.now() + 7 * 86400000).toISOString().substring(0, 10));
  const [lignes, setLignes] = useState<ILigneCapaciteCharge[]>([]);
  const [chargement, setChargement] = useState<boolean>(true);
  const [erreur, setErreur] = useState<string>("");

  const calculer = async (debut: string, fin: string): Promise<void> => {
    setChargement(true);
    setErreur("");
    try {
      const ressources: IResource[] = await resourceService.getAll();
      const resultats: ILigneCapaciteCharge[] = [];

      for (const r of ressources) {
        const [affectations, heuresAdmin] = await Promise.all([
          assignmentService.getByRessource(r.Id as number, debut, fin),
          timesheetService.getHeuresAdminParRessource(r.Id as number, debut, fin)
        ]);

        const joursOuvres = calculerJoursOuvres(debut, fin);
        const capaciteJournaliere = r.CapaciteJournaliereHeures || (r.CapaciteHebdomadaireHeures || 0) / 5;
        const capaciteTheorique = capaciteJournaliere * joursOuvres;
        const capaciteNette = resourceService.calculerCapaciteNette(capaciteTheorique, heuresAdmin);
        const chargeAffectee = affectations.reduce((s, a) => s + (a.ChargeAffecteeHeures || 0), 0);
        const tauxDeCharge = resourceService.calculerTauxDeCharge(chargeAffectee, capaciteNette);

        resultats.push({
          ressourceId: r.Id as number,
          ressourceNom: r.Title,
          capaciteTheorique,
          heuresAdmin,
          capaciteNette,
          chargeAffectee,
          tauxDeCharge,
          enSurcharge: assignmentService.detecterSurcharge(chargeAffectee, capaciteNette)
        });
      }
      setLignes(resultats);
    } catch (e) {
      setErreur("Erreur lors du calcul de la capacité/charge : " + (e as Error).message);
    } finally {
      setChargement(false);
    }
  };

  useEffect(() => {
    calculer(dateDebut, dateFin).catch(() => { /* géré */ });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const exporterCsv = (): void => {
    const entetes = ["Ressource", "Capacité théorique (h)", "Heures admin (h)", "Capacité nette (h)", "Charge affectée (h)", "Taux de charge (%)"];
    const lignesCsv = lignes.map(l => [l.ressourceNom, l.capaciteTheorique, l.heuresAdmin, l.capaciteNette, l.chargeAffectee, l.tauxDeCharge].join(";"));
    const contenu = [entetes.join(";"), ...lignesCsv].join("\n");
    const blob = new Blob([contenu], { type: "text/csv;charset=utf-8;" });
    const lien = document.createElement("a");
    lien.href = URL.createObjectURL(blob);
    lien.download = `capacite_charge_${dateDebut}_${dateFin}.csv`;
    lien.click();
  };

  const columns: IColumn[] = [
    {
      key: "alerte", name: "", minWidth: 24, maxWidth: 24,
      onRender: (i: ILigneCapaciteCharge) => i.enSurcharge
        ? <Icon iconName="Warning" title="Surcharge" style={{ color: "#a4262c" }} />
        : <Icon iconName="Accept" title="OK" style={{ color: "#107c10" }} />
    },
    { key: "ressource", name: "Ressource", fieldName: "ressourceNom", minWidth: 150, maxWidth: 200 },
    { key: "theorique", name: "Capacité théorique (h)", fieldName: "capaciteTheorique", minWidth: 130, maxWidth: 150 },
    { key: "admin", name: "Heures admin saisies (h)", fieldName: "heuresAdmin", minWidth: 140, maxWidth: 160 },
    { key: "nette", name: "Capacité nette (h)", fieldName: "capaciteNette", minWidth: 120, maxWidth: 140 },
    { key: "charge", name: "Charge affectée (h)", fieldName: "chargeAffectee", minWidth: 120, maxWidth: 140 },
    { key: "taux", name: "Taux de charge (%)", fieldName: "tauxDeCharge", minWidth: 120, maxWidth: 140 }
  ];

  return (
    <div style={biatContainerStyle}>
      <BiatHeader title="Capacité et charge" subtitle="Analyse de la charge des ressources par période" />
      <div style={biatCardStyle}>
        {erreur && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 16, borderRadius: 6 } }}>{erreur}</MessageBar>}

        <div style={biatToolbarStyle}>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="end">
            <TextField label="Du" type="date" value={dateDebut} onChange={(_, v) => setDateDebut(v || dateDebut)} />
            <TextField label="Au" type="date" value={dateFin} onChange={(_, v) => setDateFin(v || dateFin)} />
            <PrimaryButton text="Recalculer" styles={biatPrimaryBtn} onClick={() => calculer(dateDebut, dateFin).catch(() => { /* géré */ })} />
            <DefaultButton text="Exporter Excel (CSV)" styles={biatDefaultBtn} onClick={exporterCsv} />
          </Stack>
        </div>

        {chargement ? (
          <div style={{ padding: '32px 0', textAlign: 'center' }}>
            <Spinner label="Calcul en cours..." />
          </div>
        ) : (
          <DetailsList items={lignes} columns={columns} selectionMode={SelectionMode.none} setKey="ressourceId" />
        )}
      </div>
    </div>
  );
};

export default CapaciteCharge;
