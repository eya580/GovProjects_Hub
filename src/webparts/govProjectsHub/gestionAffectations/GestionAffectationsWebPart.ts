import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import GestionAffectations from "./components/GestionAffectations";
import { IGestionAffectationsProps } from "./components/IGestionAffectationsProps";

export interface IGestionAffectationsWebPartProps {
  titre: string;
}

export default class GestionAffectationsWebPart extends BaseClientSideWebPart<IGestionAffectationsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGestionAffectationsProps> = React.createElement(
      GestionAffectations,
      {
        context: this.context,
        titre: this.properties.titre || "Gestion des affectations"
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
          header: { description: "Gestion des affectations" },
          groups: [
            { groupName: "Général", groupFields: [PropertyPaneTextField("titre", { label: "Titre du web part" })] }
          ]
        }
      ]
    };
  }
}
