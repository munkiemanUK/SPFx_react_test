import * as React from 'react';
import styles from './ReactTest.module.scss';
import { IReactTestProps } from './IReactTestProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
//import { graph } from "@pnp/graph";
//import "@pnp/graph/groups";
//import "@pnp/graph/members";

//const members = graph.groups.getById("da85cb9b-8ae9-4ee9-aa51-a32d61bb08e").members();
//const owners = graph.groups.getById("da85cb9b-8ae9-4ee9-aa51-a32d61bb08e").owners();

export default class ReactTest extends React.Component<IReactTestProps, {}> {

  public render(): React.ReactElement<IReactTestProps> {

    return (
      <div className={ styles.reactTest }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <span className={styles.subTitle}>Group Members</span>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
