import * as React from "react";
import { DetailsList, IColumn, SelectionMode, Icon, TooltipHost } from "@fluentui/react";
import { IActivityNode } from "../../../../models/IActivity";
import { ActivityService } from "../../../../services/ActivityService";

export interface IComparaisonPlanningProps {
  lignes: IActivityNode[];
  activityService: ActivityService;
  indentParNiveau: number;
}

const formatDate = (d?: string): string => d ? new Date(d).toLocaleDateString() : "-";

const rendreEcart = (ecart?: number): JSX.Element => {
  if (ecart === undefined) { return <span>-</span>; }
  if (ecart === 0) {
    return <span style={{ color: "#107c10" }}><Icon iconName="Accept" /> Conforme</span>;
  }
  if (ecart > 0) {
    return (
      <TooltipHost content={`Retard de ${ecart} jour(s) par rapport au prévu`}>
        <span style={{ color: "#a4262c" }}><Icon iconName="Warning" /> +{ecart} j</span>
      </TooltipHost>
    );
  }
  return (
    <TooltipHost content={`Avance de ${Math.abs(ecart)} jour(s) par rapport au prévu`}>
      <span style={{ color: "#0078d4" }}><Icon iconName="FastForward" /> {ecart} j</span>
    </TooltipHost>
  );
};

const calculerEcartJours = (reference?: string, cible?: string): number | undefined => {
  if (!reference || !cible) { return undefined; }
  const dRef = new Date(reference);
  const dCible = new Date(cible);
  if (isNaN(dRef.getTime()) || isNaN(dCible.getTime())) { return undefined; }
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((dCible.getTime() - dRef.getTime()) / msPerDay);
};

const ComparaisonPlanning: React.FC<IComparaisonPlanningProps> = ({ lignes, activityService, indentParNiveau }) => {
  const columns: IColumn[] = [
    {
      key: "activite", name: "Activité", minWidth: 260, maxWidth: 360,
      onRender: (item: IActivityNode) => (
        <span style={{ paddingLeft: item.depth * indentParNiveau, display: "inline-block" }}>
          <strong>{item.WBS}</strong> {item.Title}
        </span>
      )
    },
    {
      key: "previsionnel", name: "Prévisionnel (début - fin)", minWidth: 160, maxWidth: 190,
      onRender: (item: IActivityNode) => `${formatDate(item.DateDebutPrevue)} → ${formatDate(item.DateFinPrevue)}`
    },
    {
      key: "reel", name: "Réel (début - fin)", minWidth: 160, maxWidth: 190,
      onRender: (item: IActivityNode) => `${formatDate(item.DateDebutReelle)} → ${formatDate(item.DateFinReelle)}`
    },
    {
      key: "ecartDebut", name: "Écart prévu vs réel (début)", minWidth: 140, maxWidth: 160,
      onRender: (item: IActivityNode) => rendreEcart(calculerEcartJours(item.DateDebutPrevue, item.DateDebutReelle))
    },
    {
      key: "ecartFin", name: "Écart prévu vs réel (fin)", minWidth: 140, maxWidth: 160,
      onRender: (item: IActivityNode) => rendreEcart(calculerEcartJours(item.DateFinPrevue, item.DateFinReelle))
    },
    {
      key: "avancement", name: "Avancement", minWidth: 80, maxWidth: 100,
      onRender: (item: IActivityNode) => `${item.PourcentageAchevement ?? 0} %`
    }
  ];

  return <DetailsList items={lignes} columns={columns} selectionMode={SelectionMode.none} setKey="Id" />;
};

export default ComparaisonPlanning;