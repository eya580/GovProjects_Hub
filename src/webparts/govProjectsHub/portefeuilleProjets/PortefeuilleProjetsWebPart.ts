import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import PortefeuilleProjets from "./components/PortefeuilleProjets";
import { IPortefeuilleProjetsProps } from "./components/IPortefeuilleProjetsProps";

export interface IPortefeuilleProjetsWebPartProps {
  titre: string;
}

export default class PortefeuilleProjetsWebPart extends BaseClientSideWebPart<IPortefeuilleProjetsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPortefeuilleProjetsProps> = React.createElement(
      PortefeuilleProjets,
      {
        context: this.context,
        titre: this.properties.titre || "Portefeuille projets"
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
          header: { description: "Portefeuille projets" },
          groups: [
            {
              groupName: "Général",
              groupFields: [
                PropertyPaneTextField("titre", { label: "Titre du web part" })
              ]
            }
          ]
        }
      ]
    };
  }
}
