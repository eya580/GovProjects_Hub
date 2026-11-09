import * as React from "react";
import { Stack, TooltipHost } from "@fluentui/react";
import { IActivityNode } from "../../../../models/IActivity";
import { StatutActivite, TypeActivite } from "../../../../models/Enums";

export interface IGanttChartProps {
  lignes: IActivityNode[];
  indentParNiveau: number;
  showLabels?: boolean;       // false = mode panneau latéral (sans colonne libellé)
  timelineWidth?: number;     // largeur dynamique de la zone barres
}

const COULEURS_STATUT: { [statut: string]: string } = {
  [StatutActivite.NonDemarree]: "#a19f9d",
  [StatutActivite.EnCours]: "#0078d4",
  [StatutActivite.Terminee]: "#107c10",
  [StatutActivite.Suspendue]: "#d83b01"
};

const HAUTEUR_LIGNE = 34;
const LARGEUR_COLONNE_LIBELLE = 320;
const LARGEUR_TIMELINE = 700;

const parseDate = (d?: string): number | undefined => (d ? new Date(d).getTime() : undefined);

const LegendeItem: React.FC<{ couleur: string; texte: string; hachure?: boolean }> = ({ couleur, texte }) => (
  <TooltipHost content={texte}>
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
      <span style={{ width: 14, height: 10, backgroundColor: couleur, display: "inline-block", borderRadius: 2 }} />
      <span style={{ fontSize: 12, color: "#323130" }}>{texte}</span>
    </Stack>
  </TooltipHost>
);

const GanttChart: React.FC<IGanttChartProps> = ({ lignes, indentParNiveau, showLabels = true, timelineWidth }) => {
  const tlWidth = timelineWidth ?? LARGEUR_TIMELINE;
  const labelWidth = showLabels ? LARGEUR_COLONNE_LIBELLE : 0;

  const bornes = React.useMemo(() => {
    const dates: number[] = [];
    lignes.forEach(l => {
      [l.DateDebutPrevue, l.DateFinPrevue, l.DateDebutReelle, l.DateFinReelle].forEach(d => {
        const t = parseDate(d);
        if (t !== undefined) { dates.push(t); }
      });
    });
    if (dates.length === 0) {
      const maintenant = Date.now();
      return { min: maintenant, max: maintenant + 30 * 86400000 };
    }
    const min = Math.min(...dates);
    const max = Math.max(...dates);
    const marge = Math.max((max - min) * 0.05, 3 * 86400000);
    return { min: min - marge, max: max + marge };
  }, [lignes]);

  const positionX = (date?: string): number | undefined => {
    const t = parseDate(date);
    if (t === undefined) { return undefined; }
    const ratio = (t - bornes.min) / (bornes.max - bornes.min || 1);
    return Math.max(0, Math.min(1, ratio)) * tlWidth;
  };

  const graduations = React.useMemo(() => {
    const nb = 6;
    const result: { x: number; label: string }[] = [];
    for (let i = 0; i <= nb; i++) {
      const t = bornes.min + ((bornes.max - bornes.min) * i) / nb;
      result.push({ x: (tlWidth * i) / nb, label: new Date(t).toLocaleDateString(undefined, { day: "2-digit", month: "2-digit" }) });
    }
    return result;
  }, [bornes, tlWidth]);

  return (
    <Stack tokens={{ childrenGap: 4 }}>
      <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { marginBottom: 8 } }}>
        <LegendeItem couleur={COULEURS_STATUT[StatutActivite.NonDemarree]} texte="Non démarrée (prévu)" />
        <LegendeItem couleur={COULEURS_STATUT[StatutActivite.EnCours]} texte="En cours (prévu)" />
        <LegendeItem couleur={COULEURS_STATUT[StatutActivite.Terminee]} texte="Terminée (prévu)" />
        <LegendeItem couleur="#323130" texte="Réel" hachure />
      </Stack>

      <div style={{ overflowX: "auto" }}>
        <svg width={labelWidth + tlWidth + 20} height={(lignes.length + 1) * HAUTEUR_LIGNE + 20}>
          {/* Graduations verticales */}
          {graduations.map((g, i) => (
            <g key={`grad-${i}`}>
              <line
                x1={labelWidth + g.x} y1={20}
                x2={labelWidth + g.x} y2={(lignes.length + 1) * HAUTEUR_LIGNE + 10}
                stroke="#edebe9" strokeWidth={1}
              />
              <text x={labelWidth + g.x} y={12} fontSize={11} fill="#605e5c" textAnchor="middle">
                {g.label}
              </text>
            </g>
          ))}

          {lignes.map((item, index) => {
            const y = 20 + index * HAUTEUR_LIGNE;
            const xDebutPrevu = positionX(item.DateDebutPrevue);
            const xFinPrevu = positionX(item.DateFinPrevue);
            const xDebutReel = positionX(item.DateDebutReelle);
            const xFinReel = positionX(item.DateFinReelle);
            const couleur = COULEURS_STATUT[item.StatutActivite] || "#a19f9d";
            const estJalon = item.EstJalon || item.TypeActivite === TypeActivite.Jalon;

            return (
              <g key={item.Id}>
                {showLabels && (
                  <text
                    x={item.depth * (indentParNiveau / 2)} y={y + HAUTEUR_LIGNE / 2 + 4}
                    fontSize={12} fill="#201f1e"
                  >
                    {"  ".repeat(item.depth)}{item.WBS} {item.Title}
                  </text>
                )}

                {/* Barre prévisionnelle */}
                {xDebutPrevu !== undefined && xFinPrevu !== undefined && !estJalon && (
                  <rect
                    x={labelWidth + xDebutPrevu}
                    y={y + 4}
                    width={Math.max(2, xFinPrevu - xDebutPrevu)}
                    height={12}
                    rx={3}
                    fill={couleur}
                    opacity={0.85}
                  />
                )}

                {/* Losange pour les jalons (dates prévues) */}
                {estJalon && xDebutPrevu !== undefined && (
                  <polygon
                    points={`${labelWidth + xDebutPrevu - 6},${y + 10} ${labelWidth + xDebutPrevu},${y + 4} ${labelWidth + xDebutPrevu + 6},${y + 10} ${labelWidth + xDebutPrevu},${y + 16}`}
                    fill={couleur}
                  />
                )}

                {/* Barre réelle (plus fine, en dessous) */}
                {xDebutReel !== undefined && xFinReel !== undefined && !estJalon && (
                  <rect
                    x={labelWidth + xDebutReel}
                    y={y + 20}
                    width={Math.max(2, xFinReel - xDebutReel)}
                    height={6}
                    rx={2}
                    fill="#323130"
                  />
                )}
              </g>
            );
          })}
        </svg>
      </div>
    </Stack>
  );
};

export default GanttChart;