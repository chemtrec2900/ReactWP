import * as React from 'react';
import styles from './ReactWebpartDemo.module.scss';
import { IReactWebpartDemoProps } from './IReactWebpartDemoProps';

import { IReactWebpartDemoState } from "./IReactWebpartDemoState";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import { IColor } from '../IColor';
import { ColorList } from './ColorList';
interface IListItem {
  Title?: string;
  Id: number;
}

export default class ReactWebpartDemo extends React.Component<IReactWebpartDemoProps, IReactWebpartDemoState> {
  private _colors: IColor[] = [
    { id: 1, title: 'red' },
    { id: 2, title: 'blue' },
    { id: 3, title: 'green' },
  ];

  constructor(props: IReactWebpartDemoProps){
    super(props);
    this.state = {colors: []};
  }

  private getColorsFromSpList(): Promise<IColor[]> {
    return new Promise<IColor[]>((resolve, reject) => {
      const endpoint: string = `${this.props.currentSiteUrl}/_api/lists/getbytitle('Colors')/items?$select=Id,Title`;
      this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          //console.log(response.headers.get("ETag"));
          return response.json();
        })
        .then((jsonResponse: any) => {
          
          jsonResponse.value.forEach(element => {
            console.log(element.Title);
          });
          let spListItemColors: IColor[] = [];
          for (let index = 0; index < jsonResponse.value.length; index++) {
            spListItemColors.push({
              id: jsonResponse.value[index].Id,
              title: jsonResponse.value[index].Title
            });
            resolve(spListItemColors);
          }
        });
    });
  }

  public componentDidMount() : void{
    this.getColorsFromSpList()
      .then((spListItemColors: IColor[]) =>{
        this.setState({colors: spListItemColors});
      });
  }

  public render(): React.ReactElement<IReactWebpartDemoProps> {
    return (
      <div className={ styles.reactWebpartDemo }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint + React!</span>
              <ColorList colors={this.state.colors} onRemoveColor = {this._removeColor}/>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _removeColor = (colorremove : IColor) : void => {
    const newColors = this.state.colors.filter(color => color != colorremove);
    this.setState({colors: newColors});
    console.log(colorremove);
    this._removeColorsFromSpList(colorremove);
  }

  private _removeColorsFromSpList(removeColor :IColor): void {
    const latestItemId = removeColor.id;
    let etag: string = undefined;
    this.props.spHttpClient.get(`${this.props.currentSiteUrl}/_api/web/lists/getbytitle('Colors')/items(${latestItemId})?$select=Id`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        etag = response.headers.get('ETag');
        return response.json();
      })
      .then((item: IListItem): Promise<SPHttpClientResponse> => {
 
    return this.props.spHttpClient.post(`${this.props.currentSiteUrl}/_api/web/lists/getbytitle('Colors')/items(${item.Id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': etag,
          'X-HTTP-Method': 'DELETE'
        }
      });
    })
    .then((response: SPHttpClientResponse): void => {
      console.log(`Item with ID: ${latestItemId} successfully deleted`);
    }, (error: any): void => {
      console.log(`Error deleting item: ${error}`);
    });
  }
}
