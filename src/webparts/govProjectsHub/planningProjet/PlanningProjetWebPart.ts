import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import PlanningProjet from "./components/PlanningProjet";
import { IPlanningProjetProps } from "./components/IPlanningProjetProps";

export interface IPlanningProjetWebPartProps {
  projetIdParDefaut: string;
}

export default class PlanningProjetWebPart extends BaseClientSideWebPart<IPlanningProjetWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPlanningProjetProps> = React.createElement(
      PlanningProjet,
      {
        context: this.context,
        projetIdParDefaut: this.properties.projetIdParDefaut ? Number(this.properties.projetIdParDefaut) : undefined
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Planning projet" },
          groups: [
            {
              groupName: "Général",
              groupFields: [
                PropertyPaneTextField("projetIdParDefaut", { label: "ID Projet par défaut (optionnel)" })
              ]
            }
          ]
        }
      ]
    };
  }
}
