import * as React from "react";
import styles from "./RenderList.module.scss";
import { IRenderListProps } from "./IRenderListProps";
import SPServices from "../Services/SPServices";

interface IRenderListState {
  listItems: any[];
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
    this.state = { listItems: [] };
    this._getListItems.bind(this);
  }

  private _getListItems() {
    this._services
      .getListItems(this.props.list, this.props.fields)
      .then((results) => {
        this.setState({ listItems: results });
        console.log(this.state.listItems);
      });
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
      </div>
    );
  }
}
