import * as React from 'react';
import styles from './RenderList.module.scss';
import { IRenderListProps } from './IRenderListProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class RenderList extends React.Component<IRenderListProps, {}> {
  public render(): React.ReactElement<IRenderListProps> {
    return (
      <div className={ styles.renderList }>
        {this.props.list}
        {
          this.props.fields && this.props.fields.map((field) => {
            return(<div>{field}</div>)
          })
        }
      </div>
    );
  }
}
