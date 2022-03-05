import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdownOptionType,
} from "@microsoft/sp-property-pane";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "RenderListWebPartStrings";
import RenderList from "./components/RenderList";
import { IRenderListProps } from "./components/IRenderListProps";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import SPServices from "./Services/SPServices";

export interface IRenderListWebPartProps {
  selectList: string;
  fields: string[];
  fieldDetails: any;
}

export default class RenderListWebPart extends BaseClientSideWebPart<IRenderListWebPartProps> {
  private _services: SPServices;
  private _lists: IPropertyPaneDropdownOption[] = [];

  private async _renderTitle() {
    this._services.getLists().then((results) => {
      results.map((list) => {
        this._lists.push({ key: list.Id, text: list.Title });
      });
    });
  }

  protected onInit(): Promise<void> {
    const sp = spfi().using(SPFx(this.context));
    this._services = new SPServices(this.context);
    this._renderTitle();
    return Promise.resolve();
  }

  public mapFields() {
    let {fields = [], fieldDetails = []} = this.properties;
    return fields.map(f => fieldDetails.find(fDetails => fDetails.key === f));
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "selectList") {
      this.properties.fieldDetails = [];
      this.properties.fields = [];
      this._services
        .getListFields(this.properties.selectList)
        .then((results) => {
          results.map((field) => {
            this.properties.fieldDetails.push({
              key: field.InternalName,
              text: field.Title,
              type: field["odata.type"],
            });
          });
          this.context.propertyPane.refresh();
        });
    }

    // this._services.getListItems(
    //   this.properties.selectList,
    //   this.mapFields()
    // );
  }

  public render(): void {
    console.log(this.properties.fields);
    const element: React.ReactElement<IRenderListProps> = React.createElement(
      RenderList,
      {
        context: this.context,
        list: this.properties.selectList,
        fields: this.mapFields(),
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
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.ConfigureList,
              groupFields: [
                PropertyPaneDropdown("selectList", {
                  label: strings.SelectListFieldLabel,
                  options: this._lists,
                }),
                PropertyFieldMultiSelect("fields", {
                  key: strings.SelectListFieldsFieldLabel,
                  label: strings.SelectListFieldsFieldLabel,
                  options: this.properties.fieldDetails,
                  selectedKeys: this.properties.fields,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
