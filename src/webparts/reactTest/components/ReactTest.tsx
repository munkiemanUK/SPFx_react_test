import * as React from 'react';
import styles from './ReactTest.module.scss';
import { IReactTestProps } from './IReactTestProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Link } from 'office-ui-fabric-react/lib/components/Link';
import { graph, GraphHttpClient } from '@pnp/graph/presets/all';
import {sp} from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import {
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClient,
  MSGraphClientFactory
} from '@microsoft/sp-http';

let groups = sp.web.currentUser.groups();
export interface IReactTestState {
  name: string;
  email: string;
  id: string;
  image: string;
}
export default class ReactTest extends React.Component<IReactTestProps, IReactTestState> {
  constructor(props: IReactTestProps) {
    super(props);
  
    this.state = {
      name: '',
      email: '',
      id: '',
      image: null
    };
  }

  public render(): React.ReactElement<IReactTestProps> {

    return (
      <div className={ styles.reactTest }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>React SPFx Testing</span>
              <p className={ styles.subTitle }>Current User</p>
              <p className={ styles.description }>Name  : {this.state.name}</p>
              <p className={ styles.description }>Email : {this.state.email}</p>
              <p className={ styles.subTitle }>365 Groups</p>
              <p className={ styles.description }>{groups}</p>
            </div>
          </div>
        </div>
      </div>
    );  
  }

  public async componentDidMount(): Promise<void> {

    this.props.graphClient
      .api('me')
      .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
        this.setState({
          name: user.displayName,
          email: user.mail,
          id: user.id,
        });
      });
  
    this.props.graphClient
      .api('/me/photo/$value')
      .responseType('blob')
      .get((err: any, photoResponse: any, rawResponse: any) => {
        const blobUrl = window.URL.createObjectURL(photoResponse);
        this.setState({ image: blobUrl });
      });

      //this._getMembers("da85cb9b-8ae9-4ee9-aa51-a32d61bb08e2");

      /*
      const queryUrl = `https://maximusunitedkingdom.sharepoint.com/_api/web/currentuser/groups`;
      const siteGroupsData = await this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1);
      const siteGroups = (await siteGroupsData.json()).value;
      siteGroups.forEach((siteGroup) => console.log(siteGroup));
      */    
  }   
}
