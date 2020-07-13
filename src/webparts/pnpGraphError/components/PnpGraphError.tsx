import * as React from 'react';
import styles from './PnpGraphError.module.scss';
import { IPnpGraphErrorProps } from './IPnpGraphErrorProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PnpGraphError extends React.Component<IPnpGraphErrorProps, {}> {
  public render(): React.ReactElement<IPnpGraphErrorProps> {
    return (
      <div className={ styles.pnpGraphError }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
