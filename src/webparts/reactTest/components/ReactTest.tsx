import * as React from 'react';
import styles from './ReactTest.module.scss';
import { IReactTestProps } from './IReactTestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { graph } from "@pnp/graph";
import '@pnp/graph/groups';

export default class ReactTest extends React.Component<IReactTestProps, {}> {
  public render(): React.ReactElement<IReactTestProps> {
    const group = graph.groups.getById("da85cb9b-8ae9-4ee9-aa51-a32d61bb08e2").expand("members")();
    console.log("group", group);

    return (
      <div className={ styles.reactTest }>
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
