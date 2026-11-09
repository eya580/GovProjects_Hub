import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import CapaciteCharge from "./components/CapaciteCharge";
import { ICapaciteChargeProps } from "./components/ICapaciteChargeProps";

export interface ICapaciteChargeWebPartProps {
  // Aucune propriété de configuration requise pour l'instant
}

export default class CapaciteChargeWebPart extends BaseClientSideWebPart<ICapaciteChargeWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICapaciteChargeProps> = React.createElement(
      CapaciteCharge,
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
        { header: { description: "Capacité et charge" }, groups: [{ groupName: "Général", groupFields: [] }] }
      ]
    };
  }
}
