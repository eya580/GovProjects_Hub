import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import MaFeuilleTemps from "./components/MaFeuilleTemps";
import { IMaFeuilleTempsProps } from "./components/IMaFeuilleTempsProps";

export interface IMaFeuilleTempsWebPartProps {
  // Aucune propriété de configuration requise pour l'instant
}

export default class MaFeuilleTempsWebPart extends BaseClientSideWebPart<IMaFeuilleTempsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMaFeuilleTempsProps> = React.createElement(
      MaFeuilleTemps,
      { context: this.context }
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
        { header: { description: "Ma feuille de temps" }, groups: [{ groupName: "Général", groupFields: [] }] }
      ]
    };
  }
}
