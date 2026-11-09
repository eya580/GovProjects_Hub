import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import Administration from "./components/Administration";
import { IAdministrationProps } from "./components/IAdministrationProps";

export interface IAdministrationWebPartProps {
  // Aucune propriété de configuration requise pour l'instant
}

export default class AdministrationWebPart extends BaseClientSideWebPart<IAdministrationWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAdministrationProps> = React.createElement(
      Administration,
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
        { header: { description: "Administration PPM" }, groups: [{ groupName: "Général", groupFields: [] }] }
      ]
    };
  }
}
