import * as React from "react";
import styles from "./RenderList.module.scss";
import { IRenderListProps } from "./IRenderListProps";
import SPServices from "../Services/SPServices";
import {
  DetailsList,
  SelectionMode,
  IColumn,
} from "@fluentui/react/lib/DetailsList";

interface IRenderListState {
  listItems: any[];
  columns: IColumn[];
}

export default class RenderList extends React.Component<
  IRenderListProps,
  IRenderListState,
  {}
> {
  private _services: SPServices;
  constructor(props: IRenderListProps) {
    super(props);
    this._services = new SPServices(this.props.context);
    this.state = { listItems: [], columns: []};
    this._getListItems.bind(this);
  }

  private _getListItems() {
    let fields = this.props.fields || [];
    this._services
      .getListItems(this.props.list, fields)
      .then((listItems) => {
        listItems = listItems && listItems.map((item) => ({
          id: item.Id, ...fields.reduce((object, field) => {
              object[field.key] = item[field.key] ? this.formatColumnValue(item[field.key], field.type) : '-';
              return object;
          }, {})
      }));
        this.setState({ listItems: listItems });
      });

      let columns: IColumn[] = [...fields].map((field) => ({
        key: field.key as string,
        name: field.text,
        minWidth: 70,
        maxWidth: 100,
        fieldName: field.key as string,
        isResizable: true,
      }))

      this.setState({columns: columns});

  }

  public formatColumnValue(value: any, type: string) {
    if (!value) {
      return value;
    }
    switch (type) {
      case 'SP.FieldDateTime':
        value = value;
        break;
      case 'SP.FieldMultiChoice':
        value = (value instanceof Array) ? value.join() : value;
        break;
      case 'SP.Taxonomy.TaxonomyField':
        value = value['Label'];
        break;
      case 'SP.FieldLookup':
        value = value['Title'];
        break;
      case 'SP.FieldUser':
        let userName = value['Title'];
        let email = value.EMail;
        value = <a href={`mailto:${email}`} style={{textDecoration: 'none', color: 'inherit'}}>{userName}</a>;
        break;
      case 'SP.FieldMultiLineText':
        value = <div dangerouslySetInnerHTML={{ __html: value }}></div>;
        break;
      case 'SP.FieldText':
        value = value;
        break;
      case 'SP.FieldComputed':
        value = value;
        break;
      case 'SP.FieldUrl':
        let url = value['Url'];
        value = url;
        break;
      case 'SP.FieldLocation':
        value = JSON.parse(value).DisplayName;
        break;
      default:
        break;
    }
    return value;
  }
  public componentDidMount(): void {
    this._getListItems();
  }

  componentDidUpdate(
    prevProps: Readonly<IRenderListProps>,
    prevState: Readonly<IRenderListState>,
    snapshot?: {}
  ): void {
    if (prevProps.fields !== this.props.fields) {
      this._getListItems();
    }
  }

  public render(): React.ReactElement<IRenderListProps> {
    return (
      <div className={styles.renderList}>
        <h2>{this.props.title}</h2>
        <DetailsList
          items={this.state.listItems}
          columns={this.state.columns}
          selectionMode= {SelectionMode.none}
        />
      </div>
    );
  }
}
