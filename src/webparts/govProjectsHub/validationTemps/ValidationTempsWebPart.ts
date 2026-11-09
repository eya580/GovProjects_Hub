import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import ValidationTemps from "./components/ValidationTemps";
import { IValidationTempsProps } from "./components/IValidationTempsProps";

export interface IValidationTempsWebPartProps {
  // Aucune propriété de configuration requise pour l'instant
}

export default class ValidationTempsWebPart extends BaseClientSideWebPart<IValidationTempsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IValidationTempsProps> = React.createElement(
      ValidationTemps,
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
        { header: { description: "Validation des feuilles de temps" }, groups: [{ groupName: "Général", groupFields: [] }] }
      ]
    };
  }
}
